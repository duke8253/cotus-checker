#!/usr/bin/env python3

# Copyright 2017 DukeGaGa
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from queue import Queue
from PIL import Image, ImageDraw, ImageFont
from gmail_secret import gmail_user, gmail_pswd
from oauth2client import tools
from google_sheets_api import get_data_from_sheet
from google_sheets_api import send_email_invalid_order
from google_sheets_api import send_email_new_order
import requests
import PyPDF2
import re
import argparse
import os
import smtplib
import json
import tempfile
import threading
import copy
import hashlib
import shutil
import time
import logging

RED = '\033[1;31m'
GREEN = '\033[1;32m'
YELLOW = '\033[1;33m'
BLUE = '\033[1;34m'
PURPLE = '\033[1;35m'
CYAN = '\033[1;36m'
WHITE = '\033[1;37m'
RESET = '\033[0;0m'

COTUS_URL = [
    'http://wwwqa.cotus.ford.com',
    'http://www.cotus.ford.com',
    'http://www.ordertracking.ford.com'
]

order_states = {
    0: 'In Order Processing:',
    1: 'In Production:',
    2: 'Awaiting Shipment:',
    3: 'In Transit:',
    4: 'Delivered:'
}

the_lock = threading.Lock()
order_str_list = []

GET_TIMEOUT = 5
GET_RETRY = 3
COTUS_RETRY = 3
COTUS_WAIT = 3

DIR_INFO = 'info'
DIR_IMAGE = 'image'
DIR_WINDOW_STICKER = 'window_sticker'

PRINT_TO_SCREEN = True


def print_to_screen(stuff_to_print):
    """
    Print stuff to the screen on demand.

    :param stuff_to_print: text to print
    :type stuff_to_print: str
    """

    if PRINT_TO_SCREEN:
        print(stuff_to_print)


def get_requests(url, payload=''):
    """
    A wrapper function to requests.get().

    :param url: the url to send the request to
    :type url: str
    :param payload: a dictionary of payload
    :type payload: dict
    :return: error number (0 for success, -1 for failure) and the response of the request
    :rtype: int, requests.api
    """

    for i in range(GET_RETRY):
        try:
            if payload:
                r = requests.get(url, params=payload, timeout=GET_TIMEOUT)
            else:
                r = requests.get(url, timeout=GET_TIMEOUT)
            return 0, r
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            pass

    return -1, '{0}SERVER TIMEOUT{1}'.format(RED, RESET)


def get_window_sticker(vin, email_addr):
    """
    Try to fetch the window sticker.

    :param vin: the vin of the car
    :type vin: str
    :param email_addr: email address is appended to the file name to distinguish files
    :type email_addr: str
    :return: error number and a message
    :rtype: int, str
    """

    # email address is appended to the file name if there is one
    if email_addr:
        file_name = os.path.join(DIR_WINDOW_STICKER, '{0}_{1}.pdf'.format(vin, email_addr))
    else:
        file_name = os.path.join(DIR_WINDOW_STICKER, '{0}.pdf'.format(vin))

    # get the hash of the old window sticker file if there is one
    hash_old = hashlib.sha256()
    sha256_old = ''
    if os.path.isfile(file_name):
        hash_old.update(open(file_name, 'rb').read())
        sha256_old = hash_old.hexdigest()

    # try fetching the window sticker
    payload = {'vin': vin}
    err, r = get_requests('http://www.windowsticker.forddirect.com/windowsticker.pdf', payload)

    # if returned error and there is NO old window sticker, return error with the response
    # if returned error and there IS an old window sticker, return success and say "FOUND BEFORE"
    if err:
        if not sha256_old:
            return -1, r
        else:
            return 0, '{0}FOUND BEFORE{1}'.format(YELLOW, RESET)

    # get a temporary name for the new window sticker and write to it
    temp_name = os.path.join(DIR_WINDOW_STICKER, '{0}.pdf'.format(next(tempfile._get_candidate_names())))
    open(temp_name, 'wb').write(r.content)

    # read the new window sticker file to and get the title of the PDF file
    with open(temp_name, 'rb') as in_file:
        pdf_reader = PyPDF2.PdfFileReader(in_file)
        pdf_title = pdf_reader.getDocumentInfo().title.lower().replace('\r', '').replace('\n', '').replace(' ', '')

    # if the title of the new PDF file says "windowsticker" after removing all other characters,
    # it means it's actually a window sticker, otherwise it's just a place holder, return "NOT FOUND"
    if pdf_title == 'windowsticker':

        # check the hash of the new window sticker
        # if it's different than the old one, then return "UPDATED"
        # if it's the same, then return "FOUND BEFORE"
        # if there is no old window sticker, then return "RELEASED"
        hash_new = hashlib.sha256()
        hash_new.update(r.content)
        sha256_new = hash_new.hexdigest()
        if sha256_old:
            if sha256_new != sha256_old:
                os.remove(file_name)
                os.rename(temp_name, file_name)
                return 2, '{0}UPDATED{1}'.format(GREEN, RESET)
            else:
                os.remove(temp_name)
                return 0, '{0}FOUND BEFORE{1}'.format(YELLOW, RESET)
        else:
            os.rename(temp_name, file_name)
            return 1, '{0}RELEASED{1}'.format(GREEN, RESET)
    else:
        os.remove(temp_name)
        return -1, '{0}NOT FOUND{1}'.format(RED, RESET)


def get_orders(file_name, new_orders=None):
    """
    Get orders from the file, combine with new orders from Google Sheet.

    :param file_name: the file name of the orders
    :type file_name: str
    :param new_orders: a list of strings of order info from Google Sheet
    :type new_orders: list[str]
    :return: a list of list of strings of the order information
    :rtype: list[list[str]]
    """

    # read the order file and put it into a list
    lines = open(file_name, 'r').readlines()

    orders = []

    # parse the order file
    for l in lines:
        o = l.replace('\n', '').replace(' ', '').split(',')
        o = list(map(str.strip, o))

        # If order uses VIN, it needs to have either 2 or 3 fields (optional email address),
        # the VIN itself must be 17 alphanumeric characters.
        #
        # If order uses Order Number and Dealer Code, it needs to have either 3 or 4 fields (optional email address),
        # Order Number must be 4 alphanumeric characters, Dealer Code must be 6 alphanumeric characters.
        #
        # Send invalid order email if any of the condition above isn't met,
        # the send_email_invalid_order will check for invalid email addresses
        if o[0] == 'vin':
            if (len(o) != 2 and len(o) != 3) or len(o[1]) != 17 or not o[1].isalnum():
                info = 'VIN, {0}'.format(', '.join(o[1:]))
                print_to_screen(info)
                print_to_screen('Invalid Order.\n')
                send_email_invalid_order(info, o[-1])
                continue
        elif o[0] == 'num':
            if (len(o) != 3 and len(o) != 4) or len(o[1]) != 4 or len(o[2]) != 6 or not o[1].isalnum() or not o[2].isalnum():
                info = 'Order Number & Dealer Code, {0}'.format(', '.join(o[1:]))
                print_to_screen(info)
                print_to_screen('Invalid Order.\n')
                send_email_invalid_order(info, o[-1])
                continue
        else:
            print_to_screen(', '.join(o))
            print_to_screen('Invalid Order.\n')
            continue

        # Make it loop pretty then put it in the order list, also makes sure no duplicates here.
        for i in range(1, len(o) - 1):
            o[i] = o[i].upper()
        o[-1] = o[-1].lower()
        o = ','.join(o)
        if o not in orders:
            orders.append(o)

    # New orders comes from google sheets, so they are already formatted,
    # just need to make sure no duplicates. Also, we print new order info to the screen.
    if new_orders is not None:
        for o in new_orders:
            if o not in orders:
                orders.append(o)
                o = o.split(',')
                if o[0] == 'vin':
                    info = 'VIN, {0}'.format(', '.join(o[1:]))
                else:
                    info = 'Order Number & Dealer Code, {0}'.format(', '.join(o[1:]))
                print_to_screen(info)
                print_to_screen('New Order.\n')
                send_email_new_order(info, o[-1])

    # Write the new orders to the order file, overwrite the old one.
    # This will guarantee the order file has no duplicates orders.
    with open(file_name, 'w') as out_file:
        for i in range(len(orders)):
            out_file.write('{0}\n'.format(orders[i]))
            orders[i] = orders[i].split(',')

    return orders


def get_data(args, which_one='', url=COTUS_URL[0]):
    """
    Get the data we need from COTUS.

    :param args: the args from argparse
    :type args: args
    :param which_one: what type of order this is
    :type which_one: str
    :param url: url to COTUS
    :type url: str
    :return: the response text
    :rtype: str
    """
    try:
        payload = {'freshLoaded': 'true'}
        if which_one == 'vin':
            payload['orderTrackingInputType'] = 'vin'
            payload['vin'] = args.vin
        else:
            payload['orderTrackingInputType'] = 'orderNumberInput'
            payload['orderNumber'] = args.order_number
            payload['dealerCode'] = args.dealer_code
            payload['customerLastName'] = args.last_name

        err, r = get_requests(url, payload)
        if err:
            return r
        else:
            return r.text.replace('\n', '').replace('\r', '')

    except KeyboardInterrupt:
        exit(2)


def get_order_info(data):
    """
    Search in the response data to find useful information.

    :param data: data returned from COTUS.
    :type data: str
    :return: order_info or error number
    :rtype: dict or int
    """

    try:

        # Use regex to search the data and put them into a dictionary
        order_info = {
            'vehicle_name': re.search(u'class="vehicleName">(.*?)</span>', data).group(1).strip(),
            'order_date': re.search(u'class="orderDate">(.*?)</span>', data).group(1).strip(),
            'order_num': re.search(u'class="orderNumber">(.*?)</span>', data).group(1).strip(),
            'dealer_code': re.search(u'"dealerInfo": { "dealerCode":(.*?)}', data).group(1).replace('"', '').strip(),
            'order_vin': re.search(u'class="vin">(.*?)</span>', data).group(1).strip(),
            'order_edd': re.search(u'id="hidden-estimated-delivery-date" data-part="(.*?)"', data).group(1).strip(),
            'current_state': re.search(u'"selectedStepName":(.*?)"surveyOn"', data).group(1).replace(',', '').replace('"', '').strip().title(),
            'email_sent': False,
            'window_sticker_sent': False,
            'initial_check_sent': False,
            'edd_changed': False,
            'state_changed': False
        }

        # some times the dealer name might not be available
        try:
            order_info['dealer_name'] = re.search(u'class="dealerName">(.*?)</span>', data).group(1).replace(',', '').replace('"', '').strip()
            if not order_info['dealer_name']:
                order_info['dealer_name'] = 'N/A'
        except AttributeError:
            order_info['dealer_name'] = 'N/A'

        # format the dates
        state_dates = [d.replace('.', '/').strip() for d in re.findall(u'Completed On : </span>(.*?)</span>', data)]
        for i in range(len(state_dates)):
            state_dates[i] = '{0}20{1}'.format(state_dates[i][:6], state_dates[i][6:])
        order_info['state_dates'] = state_dates

        # get vehicle summary
        temp = [each.strip() for each in re.findall(u'class="part-detail-description.*?>(.*?)</div>', data)]
        order_info['vehicle_summary'] = []
        for each in temp:
            if each not in order_info['vehicle_summary']:
                order_info['vehicle_summary'].append(each)

        # get the link to the rendered image of the car
        order_info['car_pic_link'] = re.search(u'http://build\.ford\.com/(?:(?!http://build\.ford\.com/|/EXT/4/vehicle\.png).)*?/EXT/4/vehicle\.png', data).group().strip()

        return order_info

    except KeyboardInterrupt:
        exit(2)

    except AttributeError:
        return -1


def get_car_image(order_info):
    """
    Generate a simple summary image using the order information.

    :param order_info: order information
    :type order_info: dict
    :return: error number
    :rtype: int
    """

    # get the image from the link, then combine the image with order information and save to a new image
    err, r = get_requests(order_info['car_pic_link'])
    if not err:
        image_file_name = os.path.join(DIR_IMAGE, '{0}.png'.format(order_info['order_vin']))
        open(image_file_name, 'wb').write(r.content)
        img = Image.open(image_file_name)
        img = img.convert('RGBA')
        width = 850 + len(order_info['vehicle_name']) * 14
        width = width if width > 1200 else 1200
        img_sig = Image.new('RGBA', (width, 359), (255, 255, 255, 255))
        img_sig.paste(img, (0, -40), img)

        fnt = ImageFont.truetype('SourceCodePro-Bold.ttf', 20)
        d = ImageDraw.Draw(img_sig)

        d.text((600, 60), 'Vehicle Name:', font=fnt, fill=(0, 0, 0))
        d.text((850, 60), order_info['vehicle_name'], font=fnt, fill=(14, 57, 201))

        d.text((600, 85), 'Ordered On:', font=fnt, fill=(0, 0, 0))
        d.text((850, 85), order_info['order_date'], font=fnt, fill=(54, 178, 8))

        d.text((600, 110), 'Estimated Delivery:', font=fnt, fill=(0, 0, 0))
        d.text((850, 110), 'N/A' if not order_info['order_edd'] else order_info['order_edd'], font=fnt,
               fill=(229, 150, 32))

        d.text((600, 135), 'Current State:', font=fnt, fill=(0, 0, 0))
        d.text((850, 135), order_info['current_state'], font=fnt, fill=(209, 6, 40))

        for i in range(5):
            d.text((600, 160 + i * 25), order_states[i], font=fnt, fill=(0, 0, 0))
            try:
                d.text((850, 160 + i * 25), 'Completed On {0}'.format(order_info['state_dates'][i]), font=fnt,
                       fill=(133, 17, 216))
            except IndexError:
                d.text((850, 160 + i * 25), 'N/A', font=fnt, fill=(133, 17, 216))

        img_sig.save(image_file_name, 'PNG')

        return 0

    return -1


def format_order_info(data, args, url):
    """
    Format the order info data into a readable string.

    :param data: the raw string data returned from get_data()
    :type data: str
    :param args: args parsed from argparse
    :type args: args
    :param url: the url used to get the data in get_data()
    :type url: str
    :return: error code, formatted str if there is no error or an error message
    :rtype: int, str
    """

    try:
        # Check if there is an error message in the data, if there is one
        # then it means the data has nothing useful (invalid order or order not found).
        error_msg = re.search(u'class="top-level-error enabled">(.*?)</p>', data).group(1).strip() + '\n'
        return -1, error_msg
    except KeyboardInterrupt:
        exit(2)
    except AttributeError:
        # If we got to here it means there is no error message in the data,
        # so we need to parse the str and put them into useful format.
        # COTUS might be unavailable from time to time, so even if there's
        # no error messages, it might just because there's nothing at all.
        order_info = get_order_info(data)
        if order_info == -1:
            return -2, 'COTUS down!'

        # Get the window sticker if needed.
        ws_err = -1
        ws_str = '{0}N/A{1}'.format(RED, RESET)
        if args.window_sticker:
            ws_err, ws_str = get_window_sticker(order_info['order_vin'], args.send_email)

        # Send email if needed.
        email_sent = '{0}N/A{1}'.format(RED, RESET)
        if args.send_email:
            email_sent = check_state(order_info, args.send_email, ws_err, args.generate_image)

        # Put the parsed data into string format so it can be printed out nicely.
        order_str = 'Order Information:\n'
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Vehicle Name:', GREEN, order_info['vehicle_name'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Ordered On:', WHITE, order_info['order_date'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Order Number:', WHITE, order_info['order_num'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Code:', WHITE, order_info['dealer_code'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('VIN:', WHITE, order_info['order_vin'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Name:', BLUE, order_info['dealer_name'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Estimated Delivery:', CYAN, 'N/A' if not order_info['order_edd'] else order_info['order_edd'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Current State:', RED, order_info['current_state'], RESET)

        # Format the dates if there are any.
        for i in range(5):
            try:
                order_str += '  {0: <21}{1}Completed On {2}{3}{4}\n'.format(order_states[i], PURPLE, GREEN, order_info['state_dates'][i], RESET)
            except IndexError:
                order_str += '  {0: <21}{1}{2}{3}\n'.format(order_states[i], PURPLE, 'N/A', RESET)

        # Where we got the information.
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Source:', YELLOW, url, RESET)

        # What happened to the window sticker.
        if args.window_sticker:
            order_str += '  {0: <21}{1}\n'.format('Window Sticker:', ws_str)

        # What happened to the email.
        if args.send_email:
            order_str += '  {0: <21}{1}\n'.format('Email Sent:', email_sent)

        # Everything about the car.
        if args.vehicle_summary:
            order_str += '  Vehicle Summary:\n'
            for each in order_info['vehicle_summary']:
                order_str += '    {0}\n'.format(each)

        # Return 1 if the status say "delivered" so we can remove this order from future checks.
        # Otherwise return 0 to say everything is fine.
        if 'delivered' in order_info['current_state'].lower():
            return 1, order_str
        else:
            return 0, order_str


def check_state(cur_data, send_email, ws_err, generate_image):
    """
    Check the previous state of the order and decide what needs to be done.

    :param cur_data: the data returned from get_order_info()
    :type cur_data: dict
    :param send_email: email address
    :type send_email: str
    :param ws_err: window sticker error code
    :type ws_err: int
    :param generate_image: whether to generate an image
    :type generate_image: bool
    :return: what happened
    :rtype: str
    """

    # File names we will be using.
    file_name = os.path.join(DIR_INFO, '{0}_{1}.json'.format(cur_data['order_vin'], send_email))
    ws_name = os.path.join(DIR_WINDOW_STICKER, '{0}_{1}.pdf'.format(cur_data['order_vin'], send_email))

    initial_check = True
    edd_changed = False
    state_changed = False
    send_ws = False
    email_sent = True

    cur_edd = cur_data['order_edd']
    cur_state = cur_data['current_state']

    # Try to read the data from previous checks if there is one.
    pre_data = None
    if os.path.isfile(file_name):
        try:
            pre_data = json.load(open(file_name, 'r'))
        except json.decoder.JSONDecodeError:
            pre_data = None

    if pre_data is not None:
        # If we found previous data it means this is not the first time,
        # so we need to do a few compares to see what we need to do.
        pre_edd = pre_data['order_edd']
        pre_state = pre_data['current_state']

        # Check if edd changed.
        if cur_edd != pre_edd:
            edd_changed = cur_data['edd_changed'] = True

        # Check if the state of the order changed.
        if cur_state != pre_state:
            state_changed = cur_data['state_changed'] = True

        # Check whether the email was sent successfully last time we checked.
        if not pre_data['email_sent']:
            email_sent = False

            # Check what changed last time and set the values accordingly.
            if not edd_changed:
                edd_changed = cur_data['edd_changed'] = pre_data['edd_changed']
            if not state_changed:
                state_changed = cur_data['state_changed'] = pre_data['state_changed']

        # Check if we have a window sticker file, and if it were sent before.
        if os.path.isfile(ws_name):
            send_ws = not pre_data['window_sticker_sent']

            # If the window sticker wasn't sent before, we need to send it now.
            # We also need to send the window sticker if it was updated (ws_err == 2), removed for now.
            # Now we only send two copies of the window sticker, when it was first released,
            # and when the car is delivered.
            if send_ws:
                ws_err = 1
            elif 'delivered' in cur_data['current_state'].lower():
                send_ws = True
                ws_err = 3
            # elif ws_err == 2:
            #     send_ws = True

        # A few carry over flags.
        initial_check = not pre_data['initial_check_sent']
        cur_data['email_sent'] = pre_data['email_sent']
        cur_data['initial_check_sent'] = pre_data['initial_check_sent']
        cur_data['window_sticker_sent'] = pre_data['window_sticker_sent']
    else:
        # If it's the first time, everything needs to be sent.
        if cur_edd:
            edd_changed = cur_data['edd_changed'] = True

        state_changed = cur_data['state_changed'] = True

        if os.path.isfile(ws_name):
            send_ws = True
            ws_err = 1

    # Send email if something changed, or the previous attempt failed.
    err = -1
    ret_msg = '{0}STATUS NOT CHANGED{1}'.format(YELLOW, RESET)
    if edd_changed or state_changed or send_ws or not email_sent:

        # Generate the image if needed.
        img_err = -1
        if generate_image:
            img_err = get_car_image(cur_data)

        # What happened to the EDD.
        if edd_changed:
            if not cur_edd:
                edd = 'Removed'
            else:
                edd = cur_edd
        else:
            edd = ''

        # Send email.
        err, ret_msg = report_with_email(
            send_email,
            edd,
            cur_state if state_changed else '',
            cur_data['order_vin'],
            initial_check,
            send_ws,
            ws_err,
            img_err
        )

    # Update a few flags if everything went through.
    if not err:
        cur_data['email_sent'] = True
        cur_data['initial_check_sent'] = True
        if send_ws:
            cur_data['window_sticker_sent'] = True

    # Save the new status of the order to the file, overwriting the old one.
    json.dump(cur_data, open(file_name, 'w'), indent=2)

    return ret_msg


def report_with_email(email_to, edd='', state='', vin='', initial_check=False, send_ws=False, ws_err=0, img_err=-1):
    """
    Send the email.

    :param email_to: the email address to send to
    :type email_to: str
    :param edd: the EDD of the order
    :type edd: str
    :param state: the current state of the order
    :type state: str
    :param vin: the VIN
    :type vin: str
    :param initial_check: whether it's the first time
    :type initial_check: bool
    :param send_ws: whether to send window sticker
    :type send_ws: bool
    :param ws_err: error code returned by get_window_sticker()
    :type ws_err: int
    :param img_err: error code returned by get_car_image()
    :type img_err: int
    :return: error code, error message
    :rtype: int, str
    """

    if not gmail_user or not gmail_pswd:
        # We need user name and password to send emails.
        return -1, '{0}Empty Gmail Username or Password{1}'.format(RED, RESET)
    else:
        email_from = gmail_user
        email_body = ''

        # Format the email body, standard stuff.
        if edd:
            email_body += 'EDD: {0}\n'.format(edd)
        if state:
            email_body += 'Status: {0}\n'.format(state.title())
        if send_ws:
            if ws_err == 1:
                email_body += 'Window Sticker Released!\n'
            elif ws_err == 2 or ws_err == 3:
                email_body += 'Window Sticker Updated!\n'

        email_msg = MIMEMultipart()
        if initial_check:
            email_msg['Subject'] = '[COTUS CHECKER] Order Status Changed for VIN: {0} (Initial Check)'.format(vin)
        else:
            email_msg['Subject'] = '[COTUS CHECKER] Order Status Changed for VIN: {0}'.format(vin)
        email_msg['From'] = gmail_user
        email_msg['To'] = email_to
        email_msg['Date'] = formatdate(localtime=True)
        email_msg.attach(MIMEText(email_body))

        # Attach the window sticker file to the email if needed.
        if send_ws:
            attachment_name = '{0}.pdf'.format(vin)
            file_name = os.path.join(DIR_WINDOW_STICKER, '{0}_{1}.pdf'.format(vin, email_to))
            attachment = MIMEApplication(open(file_name, 'rb').read())
            attachment.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(attachment_name))
            email_msg.attach(attachment)

        # Attach the image file to the email if needed.
        if not img_err:
            attachment_name = '{0}.png'.format(vin)
            file_name = os.path.join(DIR_IMAGE, attachment_name)
            attachment = MIMEApplication(open(file_name, 'rb').read())
            attachment.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(attachment_name))
            email_msg.attach(attachment)

        # Try to send the email, return 0 on success, -1 on fail.
        try:
            gmail_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            gmail_server.ehlo()
            gmail_server.login(gmail_user, gmail_pswd)
            gmail_server.sendmail(email_from, email_to, email_msg.as_string())
            gmail_server.close()
            return 0, '{0}SUCCESS{1}'.format(GREEN, RESET)
        except KeyboardInterrupt:
            exit(2)
        except (smtplib.SMTPException, smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException,
                smtplib.SMTPSenderRefused, smtplib.SMTPRecipientsRefused, smtplib.SMTPDataError, smtplib.SMTPConnectError,
                smtplib.SMTPHeloError, smtplib.SMTPNotSupportedError, smtplib.SMTPAuthenticationError):
            return -1, '{0}FAIL{1}'.format(RED, RESET)


def check_order(q_in, q_out, q_count):
    """
    Thread.

    :param q_in: input data from the order file, and other related info
    :type q_in: Queue
    :param q_out: keep track of which orders needs to be removed from the order file
    :type q_out: Queue
    :param q_count: keep track of how many orders each thread checked successfully
    :type q_count: Queue
    """

    global order_str_list

    # Keep track of how many orders each thread checked successfully.
    count = 0

    # Keep checking until the input queue is empty.
    while not q_in.empty():

        # Get the info needed to check an order, and reset the stop flag.
        args, order, list_id = q_in.get()
        err = -1
        msg = ''
        stop_flag = False

        # Keep retrying until hitting the retry limit, or until one check succeeds.
        for i in range(len(COTUS_URL)):
            if stop_flag:
                break
            url = COTUS_URL[i]
            for j in range(COTUS_RETRY):
                if stop_flag:
                    break

                # Get the data then format it.
                data = get_data(args, order[0], url=url)
                err, msg = format_order_info(data, args, url)

                # Stop trying if nothing went wrong.
                if err >= 0:
                    stop_flag = True
                    count += 1
                else:
                    time.sleep(COTUS_WAIT)

        if err == 1:
            # Put the index of the current order into the out queue so it'll be removed.
            q_out.put(list_id)
        elif err == -1:
            # Format the error message.
            if order[0] == 'vin':
                if len(order) == 2:
                    msg = 'VIN: {0}\n{1}'.format(order[1], msg)
                else:
                    msg = 'VIN: {0}, Email: {1}\n{2}'.format(order[1], order[2], msg)
            else:
                if len(order) == 3:
                    msg = 'Order Number: {0}, Dealer Code: {1}\n{2}'.format(order[1], order[2], msg)
                else:
                    msg = 'Order Number: {0}, Dealer Code: {1}, Email: {2}\n{3}'.format(order[1], order[2], order[3], msg)

        # Put the message into the global list using the index so
        # it can be printed out in the same order of the order file.
        the_lock.acquire()
        order_str_list[list_id] = msg
        the_lock.release()

    # Put the total number of orders checked by this thread in the queue.
    q_count.put(count)


def main():
    """
    The main function.

    :return: error number
    :rtype: int
    """
    global order_str_list, DIR_INFO, DIR_IMAGE, DIR_WINDOW_STICKER, PRINT_TO_SCREEN

    # Get the path of the file, extract the directory path from it, and set the work directory to it.
    my_abspath = os.path.abspath(__file__)
    my_dirname = os.path.dirname(my_abspath)
    os.chdir(my_dirname)

    # Setup all the other directories needed, create them if necessary.
    DIR_INFO = os.path.join(my_dirname, DIR_INFO)
    DIR_IMAGE = os.path.join(my_dirname, DIR_IMAGE)
    DIR_WINDOW_STICKER = os.path.join(my_dirname, DIR_WINDOW_STICKER)
    if not os.path.isdir(DIR_INFO):
        shutil.rmtree(DIR_INFO, ignore_errors=True)
        os.mkdir(DIR_INFO)
    if not os.path.isdir(DIR_IMAGE):
        shutil.rmtree(DIR_IMAGE, ignore_errors=True)
        os.mkdir(DIR_IMAGE)
    if not os.path.isdir(DIR_WINDOW_STICKER):
        shutil.rmtree(DIR_WINDOW_STICKER, ignore_errors=True)
        os.mkdir(DIR_WINDOW_STICKER)

    # setup the arguments
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument('-o', '--order-number', type=str, help='order number of the car', dest='order_number')
    parser.add_argument('-d', '--dealer-code', type=str, help='dealer code of the order', dest='dealer_code')
    parser.add_argument('-l', '--last-name', type=str, help='customer\'s last name (not used for now)', dest='last_name', default='xxx')
    parser.add_argument('-v', '--vin', type=str, help='VIN of the car', dest='vin')
    parser.add_argument('-s', '--vehicle-summary', help='show vehicle summary', dest='vehicle_summary', action='store_true', default=False)
    parser.add_argument('-f', '--file', type=str, help='file with many many VIN\'s', dest='file')
    parser.add_argument('-e', '--send-email', type=str, help='send email if state changed', dest='send_email')
    parser.add_argument('-w', '--window-sticker', help='obtain the window sticker', dest='window_sticker', action='store_true', default=False)
    parser.add_argument('-r', '--remove-delivered', help='remove delivered orders from the file', dest='remove_delivered', action='store_true', default=False)
    parser.add_argument('-i', '--generate-image', help='generate an image with the dates and the car on it', dest='generate_image', action='store_true', default=False)
    parser.add_argument('-n', '--no-print', help='print stuff to the screen', dest='no_print', action='store_true', default=False)
    args = parser.parse_args()

    PRINT_TO_SCREEN = not args.no_print

    if args.file:
        if not os.path.isfile(args.file):
            print_to_screen('Invalid VIN file.')
            exit(1)
        else:

            # Since all the responses will be printed on the screen,
            # we only log the process if there's an input order file given,
            # no sense to log only one query.
            logger = logging.getLogger('COTUS Checker')
            logger.setLevel(logging.DEBUG)
            log_handler = logging.FileHandler('logs.log')
            log_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            log_handler.setLevel(logging.DEBUG)
            log_handler.setFormatter(log_formatter)
            logger.addHandler(log_handler)

            # create queues for passing data between threads
            q_in = Queue()
            q_out = Queue()
            q_count = Queue()

            # Create 10 threads, don't want to stress the server too much, it's not a DDoS.
            threads = [threading.Thread(target=check_order, args=(q_in, q_out, q_count)) for i in range(10)]

            # Fetch new orders from google sheets, and read in from the order file.
            orders = get_orders(args.file, get_data_from_sheet(args, my_dirname))

            # Different order types needs different data.
            # Just being lazy here, using existing args variable since we don't need the other parts of it.
            for o in orders:
                if o[0] == 'vin':
                    if len(o) == 2:
                        args.order_number = ''
                        args.dealer_code = ''
                        args.vin = o[1]
                        args.send_email = ''
                    else:
                        args.order_number = ''
                        args.dealer_code = ''
                        args.vin = o[1]
                        args.send_email = o[2]
                else:
                    if len(o) == 3:
                        args.order_number = o[1]
                        args.dealer_code = o[2]
                        args.vin = ''
                        args.send_email = ''
                    else:
                        args.order_number = o[1]
                        args.dealer_code = o[2]
                        args.vin = ''
                        args.send_email = o[3]

                # Must make copy so we can have put them into different threads without messing up.
                # The length of the order_str_list will be the index to write to it.
                q_in.put((copy.deepcopy(args), o, len(order_str_list)))
                order_str_list.append('')

            # Start all threads and wait for them to finish.
            for t in threads:
                t.start()
            for t in threads:
                t.join()

            # Print out all the query results.
            for s in order_str_list:
                print_to_screen(s)

            # Log what we did just now.
            logger.info('Total Orders: {0}, Query Success: {1}'.format(len(orders), sum(list(q_count.queue))))

            # Remove orders that are marked "Delivered" from the order file.
            if args.remove_delivered:
                remove_list = list(q_out.queue)
                with open(args.file, 'w') as out_file:
                    for i in range(len(orders)):
                        if i not in remove_list:
                            out_file.write('{0}\n'.format(','.join(orders[i])))

    else:

        # if we're not using a order file
        data = None
        msg = ''
        stop_flag = False
        for i in range(len(COTUS_URL)):
            if stop_flag:
                break
            url = COTUS_URL[i]
            for j in range(COTUS_RETRY):
                if stop_flag:
                    break
                if args.vin:
                    data = get_data(args, 'vin', url=url)
                elif args.order_number and args.dealer_code and args.last_name:
                    data = get_data(args, url=url)
                else:
                    print_to_screen('Invalid input!')
                    exit(1)
                err, msg = format_order_info(data, args, url)
                if err >= 0:
                    stop_flag = True
                else:
                    time.sleep(COTUS_WAIT)
        print_to_screen(msg)


if __name__ == '__main__':
    main()
