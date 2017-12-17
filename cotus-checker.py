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
import requests
import textract
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

DIR_INFO = 'info'
DIR_IMAGE = 'image'
DIR_WINDOW_STICKER = 'window_sticker'


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
        except requests.exceptions.Timeout:
            pass
    return -1, '{0}SERVER TIMEOUT{1}'.format(RED, RESET)


def get_window_sticker(vin):
    """
    Try to fetch the window sticker.

    :param vin: the vin of the car
    :type vin: str
    :return: error number and a message
    :rtype: int, str
    """
    file_name = os.path.join(DIR_WINDOW_STICKER, '{0}.pdf'.format(vin))
    hash_old = hashlib.sha256()
    sha256_old = ''

    if os.path.isfile(file_name):
        with open(file_name, 'rb') as in_file:
            hash_old.update(in_file.read())
        sha256_old = hash_old.hexdigest()

    temp_name = os.path.join(DIR_WINDOW_STICKER, '{0}.pdf'.format(next(tempfile._get_candidate_names())))
    payload = {'vin': vin}
    err, r = get_requests('http://www.windowsticker.forddirect.com/windowsticker.pdf', payload)
    if err:
        if not sha256_old:
            return -1, r
        else:
            return 0, '{0}FOUND BEFORE{1}'.format(YELLOW, RESET)

    open(temp_name, 'wb').write(r.content)
    text = textract.process(temp_name).decode('utf-8')

    if 'BLEND' in text:
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


def get_orders(file_name, new_orders):
    """
    Get orders from the file, combine with new orders from Google Sheet.

    :param file_name: the file name of the orders
    :type file_name: str
    :param new_orders: a list of strings of order info from Google Sheet
    :type new_orders: list[str]
    :return: a list of list of strings of the order information
    :rtype: list[list[str]]
    """
    with open(file_name, 'r') as in_file:
        lines = in_file.readlines()

    orders = []
    for l in lines:
        o = l.replace('\n', '').replace(' ', '').split(',')
        o = list(map(str.strip, o))
        if o[0] == 'vin':
            if (len(o) != 2 and len(o) != 3) or len(o[1]) != 17 or not o[1].isalnum():
                info = 'VIN, {0}'.format(', '.join(o[1:]))
                print(info)
                print('Invalid Order.\n')
                send_email_invalid_order(info, o[-1])
                continue
        elif o[0] == 'num':
            if (len(o) != 3 and len(o) != 4) or len(o[1]) != 4 or len(o[2]) != 6 or not o[1].isalnum() or not o[2].isalnum():
                info = 'Order Number & Dealer Code, {0}'.format(', '.join(o[1:]))
                print(info)
                print('Invalid Order.\n')
                send_email_invalid_order(info, o[-1])
                continue
        else:
            print(', '.join(o))
            print('Invalid Order.\n')
            continue
        for i in range(1, len(o) - 1):
            o[i] = o[i].upper()
        o[-1] = o[-1].lower()
        o = ','.join(o)
        if o not in orders:
            orders.append(o)

    for o in new_orders:
        if o not in orders:
            orders.append(o)

    with open(file_name, 'w') as out_file:
        for i in range(len(orders)):
            out_file.write('{0}\n'.format(orders[i]))
            orders[i] = orders[i].split(',')

    return orders


def get_data(args, which_one='', url=COTUS_URL[0]):
    """

    :param args:
    :type args:
    :param which_one:
    :type which_one:
    :param url:
    :type url:
    :return:
    :rtype:
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

    :param data:
    :type data:
    :return:
    :rtype:
    """
    try:
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

        try:
            order_info['dealer_name'] = re.search(u'class="dealerName">(.*?)</span>', data).group(1).replace(',', '').replace('"', '').strip()
            if not order_info['dealer_name']:
                order_info['dealer_name'] = 'N/A'
        except AttributeError:
            order_info['dealer_name'] = 'N/A'

        state_dates = [d.replace('.', '/').strip() for d in re.findall(u'Completed On : </span>(.*?)</span>', data)]
        for i in range(len(state_dates)):
            if state_dates[i][5:] == '/17':
                state_dates[i] = state_dates[i][:6] + '2017'
            elif state_dates[i][5:] == '/18':
                state_dates[i] = state_dates[i][:6] + '2018'
        order_info['state_dates'] = state_dates

        temp = [each.strip() for each in re.findall(u'class="part-detail-description.*?>(.*?)</div>', data)]
        order_info['vehicle_summary'] = []
        for each in temp:
            if each not in order_info['vehicle_summary']:
                order_info['vehicle_summary'].append(each)

        order_info['car_pic_link'] = re.search(u'http://build\.ford\.com/(?:(?!http://build\.ford\.com/|/EXT/4/vehicle\.png).)*?/EXT/4/vehicle\.png', data).group().strip()

        return order_info
    except KeyboardInterrupt:
        exit(2)
    except AttributeError:
        return -1


def get_car_image(order_info):
    """
    Generate a simple summary image using the order information.

    :param order_info:
    :type order_info:
    :return:
    :rtype:
    """
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


def format_order_info(data, vehicle_summary=False, send_email='', url=COTUS_URL[0], window_sticker=False, generate_image=False):
    """
    Format the order info data into a readable string.

    :param data:
    :type data:
    :param vehicle_summary:
    :type vehicle_summary:
    :param send_email:
    :type send_email:
    :param url:
    :type url:
    :param window_sticker:
    :type window_sticker:
    :param generate_image:
    :type generate_image:
    :return:
    :rtype:
    """
    try:
        error_msg = re.search(u'class="top-level-error enabled">(.*?)</p>', data).group(1).strip() + '\n'
        return -1, error_msg
    except KeyboardInterrupt:
        exit(2)
    except AttributeError:
        order_info = get_order_info(data)
        if order_info == -1:
            return -2, 'COTUS down!'

        ws_err = -1
        ws_str = '{0}N/A{1}'.format(RED, RESET)
        if window_sticker:
            ws_err, ws_str = get_window_sticker(order_info['order_vin'])

        email_sent = '{0}N/A{1}'.format(RED, RESET)
        if send_email:
            email_sent = check_state(order_info, send_email, ws_err, generate_image)

        order_str = 'Order Information:\n'
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Vehicle Name:', GREEN, order_info['vehicle_name'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Ordered On:', WHITE, order_info['order_date'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Order Number:', WHITE, order_info['order_num'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Code:', WHITE, order_info['dealer_code'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('VIN:', WHITE, order_info['order_vin'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Name:', BLUE, order_info['dealer_name'], RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Estimated Delivery:', CYAN,
                                                    'N/A' if not order_info['order_edd'] else order_info['order_edd'],
                                                    RESET)
        order_str += '  {0: <21}{1}{2}{3}\n'.format('Current State:', RED, order_info['current_state'], RESET)

        for i in range(5):
            try:
                order_str += '  {0: <21}{1}Completed On {2}{3}{4}\n'.format(order_states[i], PURPLE, GREEN,
                                                                            order_info['state_dates'][i], RESET)
            except IndexError:
                order_str += '  {0: <21}{1}{2}{3}\n'.format(order_states[i], PURPLE, 'N/A', RESET)

        order_str += '  {0: <21}{1}{2}{3}\n'.format('Source:', YELLOW, url, RESET)

        if window_sticker:
            order_str += '  {0: <21}{1}\n'.format('Window Sticker:', ws_str)

        if send_email:
            order_str += '  {0: <21}{1}\n'.format('Email Sent:', email_sent)

        if vehicle_summary:
            order_str += '  Vehicle Summary:\n'
            for each in order_info['vehicle_summary']:
                order_str += '    {0}\n'.format(each)

        if 'delivered' in order_info['current_state'].lower():
            return 1, order_str
        else:
            return 0, order_str


def check_state(cur_data, send_email, ws_err, generate_image):
    """

    :param cur_data:
    :type cur_data:
    :param send_email:
    :type send_email:
    :param ws_err:
    :type ws_err:
    :param generate_image:
    :type generate_image:
    :return:
    :rtype:
    """
    file_name = os.path.join(DIR_INFO, '{0}.json'.format(cur_data['order_vin']))
    ws_name = os.path.join(DIR_WINDOW_STICKER, '{0}.pdf'.format(cur_data['order_vin']))

    initial_check = True
    edd_changed = False
    state_changed = False
    send_ws = False
    email_sent = True

    cur_edd = cur_data['order_edd']
    cur_state = cur_data['current_state']

    pre_data = None
    if os.path.isfile(file_name):
        pre_data = json.load(open(file_name, 'r'))

    if pre_data is not None:
        pre_edd = pre_data['order_edd']
        pre_state = pre_data['current_state']

        if cur_edd != pre_edd:
            edd_changed = cur_data['edd_changed'] = True

        if cur_state != pre_state:
            state_changed = cur_data['state_changed'] = True

        if not pre_data['email_sent']:
            email_sent = False
            edd_changed = cur_data['edd_changed'] = pre_data['edd_changed']
            state_changed = cur_data['state_changed'] = pre_data['state_changed']

        if os.path.isfile(ws_name):
            send_ws = not pre_data['window_sticker_sent']
            if send_ws:
                ws_err = 1
            elif ws_err == 2:
                send_ws = True

        initial_check = not pre_data['initial_check_sent']
        cur_data['email_sent'] = pre_data['email_sent']
        cur_data['initial_check_sent'] = pre_data['initial_check_sent']
        cur_data['window_sticker_sent'] = pre_data['window_sticker_sent']
    else:
        if cur_edd:
            edd_changed = cur_data['edd_changed'] = True

        state_changed = cur_data['state_changed'] = True

        if os.path.isfile(ws_name):
            send_ws = True
            ws_err = 1

    err = -1
    ret_msg = '{0}STATUS NOT CHANGED{1}'.format(YELLOW, RESET)
    if edd_changed or state_changed or send_ws or not email_sent:
        img_err = -1
        if generate_image:
            img_err = get_car_image(cur_data)

        if edd_changed:
            if not cur_edd:
                edd = 'Removed'
            else:
                edd = cur_edd
        else:
            edd = ''

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

    if not err:
        cur_data['email_sent'] = True
        cur_data['initial_check_sent'] = True
        if send_ws:
            cur_data['window_sticker_sent'] = True

    json.dump(cur_data, open(file_name, 'w'), indent=2)

    return ret_msg


def report_with_email(email_to, edd='', state='', vin='', initial_check=False, send_ws=False, ws_err=0, img_err=-1):
    """

    :param email_to:
    :type email_to:
    :param edd:
    :type edd:
    :param state:
    :type state:
    :param vin:
    :type vin:
    :param initial_check:
    :type initial_check:
    :param send_ws:
    :type send_ws:
    :param ws_err:
    :type ws_err:
    :param img_err:
    :type img_err:
    :return:
    :rtype:
    """
    if not gmail_user or not gmail_pswd:
        return -1, '{0}Empty Gmail Username or Password{1}'.format(RED, RESET)
    else:
        email_from = gmail_user
        email_body = ''
        if edd:
            email_body += 'EDD: {0}\n'.format(edd)
        if state:
            email_body += 'Status: {0}\n'.format(state.title())
        if send_ws:
            if ws_err == 1:
                email_body += 'Window Sticker Released!\n'
            else:
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

        if send_ws:
            base_name = '{0}.pdf'.format(vin)
            file_name = os.path.join(DIR_WINDOW_STICKER, base_name)
            with open(file_name, 'rb') as in_file:
                attachment = MIMEApplication(in_file.read(), Name=base_name)
            attachment['Content-Disposition'] = 'attachment; filename="{0}"'.format(base_name)
            email_msg.attach(attachment)

        if not img_err:
            base_name = '{0}.png'.format(vin)
            file_name = os.path.join(DIR_IMAGE, base_name)
            with open(file_name, 'rb') as in_file:
                attachment = MIMEApplication(in_file.read(), Name=base_name)
            attachment['Content-Disposition'] = 'attachment; filename="{0}"'.format(base_name)
            email_msg.attach(attachment)

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


def check_order(q_in, q_out):
    """

    :param q_in:
    :type q_in:
    :param q_out:
    :type q_out:
    :return:
    :rtype:
    """
    global order_str_list

    while not q_in.empty():
        args, order, list_id = q_in.get()
        err = -1
        msg = ''
        for i in range(len(COTUS_URL)):
            url = COTUS_URL[i]
            data = get_data(args, order[0], url=url)
            err, msg = format_order_info(data, args.vehicle_summary, args.send_email, url, args.window_sticker, args.generate_image)
            if err >= 0:
                break
        if err == 1:
            q_out.put(list_id)
        elif err == -1:
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
        the_lock.acquire()
        order_str_list[list_id] = msg
        the_lock.release()


def main():
    """

    :return:
    :rtype:
    """
    global order_str_list, DIR_INFO, DIR_IMAGE, DIR_WINDOW_STICKER

    my_abspath = os.path.abspath(__file__)
    my_dirname = os.path.dirname(my_abspath)
    os.chdir(my_dirname)

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
    args = parser.parse_args()

    if args.file:
        if not os.path.isfile(args.file):
            print('Invalid VIN file.')
            exit(1)
        else:
            q_in = Queue()
            q_out = Queue()
            threads = [threading.Thread(target=check_order, args=(q_in, q_out)) for i in range(10)]

            orders = get_orders(args.file, get_data_from_sheet(args, my_dirname))
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

                q_in.put((copy.deepcopy(args), o, len(order_str_list)))
                order_str_list.append('')
            for t in threads:
                t.start()
            for t in threads:
                t.join()
            for s in order_str_list:
                print(s)

            if args.remove_delivered:
                remove_list = list(q_out.queue)
                with open(args.file, 'w') as out_file:
                    for i in range(len(orders)):
                        if i not in remove_list:
                            out_file.write('{0}\n'.format(','.join(orders[i])))

    else:
        data = None
        msg = ''
        for i in range(len(COTUS_URL)):
            url = COTUS_URL[i]
            if args.vin:
                data = get_data(args, 'vin', url=url)
            elif args.order_number and args.dealer_code and args.last_name:
                data = get_data(args, url=url)
            else:
                print('Invalid input!')
                exit(1)
            err, msg = format_order_info(data, args.vehicle_summary, args.send_email, url, args.window_sticker, args.generate_image)
            if err >= 0:
                break
        print(msg)


if __name__ == '__main__':
    main()
