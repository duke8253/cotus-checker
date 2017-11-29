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

RED    = '\033[1;31m'
GREEN  = '\033[1;32m'
YELLOW = '\033[1;33m'
BLUE   = '\033[1;34m'
PURPLE = '\033[1;35m'
CYAN   = '\033[1;36m'
WHITE  = '\033[1;37m'
RESET  = '\033[0;0m'

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

# Set your own gmail address and password
gmail_user = ''
gmail_pswd = ''

def get_window_sticker(vin):
  file_name = '{0}.pdf'.format(vin)
  if os.path.isfile(file_name):
    return '{0}FOUND BEFORE{1}'.format(YELLOW, RESET)

  temp_name = '{0}.pdf'.format(next(tempfile._get_candidate_names()))
  payload = {'vin': vin}
  try:
    r = requests.get('http://www.windowsticker.forddirect.com/windowsticker.pdf', params=payload, timeout=3)
  except requests.exceptions.Timeout:
    return '{0}SERVER TIMEOUT{1}'.format(RED, RESET)

  open(temp_name, 'wb').write(r.content)
  text = textract.process(temp_name).decode('utf-8')
  if 'BLEND' in text:
    os.rename(temp_name, file_name)
    return '{0}FOUND{1}'.format(GREEN, RESET)

  os.remove(temp_name)
  return '{0}NOT FOUND{1}'.format(RED, RESET)

def get_orders(file_name):
  with open(file_name, 'r') as in_file:
    lines = in_file.readlines()
  orders = []
  for l in lines:
    orders.append(l.replace('\n', '').strip().split(','))
  return orders

def get_data(args, which_one='', url=COTUS_URL[0]):
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
    try:
      r = requests.get(url, params=payload, timeout=3)
      return r.text.replace('\n', '').replace('\r', '')
    except requests.exceptions.Timeout:
      return 'SERVER TIMEOUT'

  except KeyboardInterrupt:
    exit(2)
  except:
    return ''

def get_order_info(data):
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
      'initial_check_sent': False
    }

    try:
      r = requests.get('http://www.fleet.ford.com/fleetdealers', params={'dealerCode': order_info['dealer_code']}, timeout=3)
      order_info['dealer_name'] = re.search(u'class="dealer-name">.*?>(.*?)</a>', r.text.replace('\n', '').replace('\r', '')).group(1).strip()
    except (AttributeError, requests.exceptions.Timeout):
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

    return order_info
  except KeyboardInterrupt:
    exit(2)
  except AttributeError:
    return -1

def format_order_info(data, vehicle_summary=False, send_email='', url=COTUS_URL[0], window_sticker=False):
  try:
    error_msg = re.search(u'class="top-level-error enabled">(.*?)</p>', data).group(1).strip() + '\n'
    return 1, error_msg
  except KeyboardInterrupt:
    exit(2)
  except AttributeError:
    order_info = get_order_info(data)
    if order_info == -1:
      return 2, 'COTUS down!'

    if window_sticker:
      ws_str = get_window_sticker(order_info['order_vin'])

    if send_email:
      email_sent = check_state(order_info, send_email)

    order_str = 'Order Information:\n'
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Vehicle Name:', GREEN, order_info['vehicle_name'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Ordered On:', WHITE, order_info['order_date'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Order Number:', WHITE, order_info['order_num'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Code:', WHITE, order_info['dealer_code'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('VIN:', WHITE, order_info['order_vin'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Dealer Name:', BLUE, order_info['dealer_name'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Estimated Delivery:', CYAN, 'N/A' if not order_info['order_edd'] else order_info['order_edd'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Current State:', RED, order_info['current_state'], RESET)

    for i in range(5):
      try:
        order_str += '  {0: <21}{1}{2}{3}\n'.format(order_states[i], PURPLE, order_info['state_dates'][i], RESET)
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

    return 0, order_str

def check_state(cur_data, send_email):
  file_name = '{0}.txt'.format(cur_data['order_vin'])
  ws_name = '{0}.pdf'.format(cur_data['order_vin'])

  initial_check = True
  edd_changed = False
  state_changed = False
  send_ws = False

  cur_edd = cur_data['order_edd']
  cur_state = cur_data['current_state']

  pre_data = None
  if os.path.isfile(file_name):
    pre_data = json.load(open(file_name, 'r'))

  if pre_data is not None:
    pre_edd = pre_data['order_edd']
    pre_state = pre_data['current_state']

    if cur_edd and (cur_edd != pre_edd or not pre_data['email_sent']):
      edd_changed = True

    if cur_state != pre_state or not pre_data['email_sent']:
      state_changed = True

    if not edd_changed and not state_changed:
      cur_data['email_sent'] = pre_data['email_sent']

    if os.path.isfile(ws_name):
      send_ws = not pre_data['window_sticker_sent']

    initial_check = not pre_data['initial_check_sent']
    cur_data['initial_check_sent'] = pre_data['initial_check_sent']
    cur_data['window_sticker_sent'] = pre_data['window_sticker_sent']
  else:
    if cur_edd:
      edd_changed = True

    state_changed = True

    if os.path.isfile(ws_name):
      send_ws = True

  err = -1
  email_sent = '{0}STATUS NOT CHANGED{1}'.format(YELLOW, RESET)
  if edd_changed or state_changed or send_ws:
    err, email_sent = report_with_email(
      send_email,
      cur_edd if edd_changed else '',
      cur_state if state_changed else '',
      cur_data['order_vin'],
      initial_check,
      send_ws
    )

  if not err:
    cur_data['email_sent'] = True
    cur_data['initial_check_sent'] = True
    if send_ws:
      cur_data['window_sticker_sent'] = True

  json.dump(cur_data, open(file_name, 'w'), indent=2)

  return email_sent

def report_with_email(email_to, edd='', state='', vin='', initial_check=False, send_ws=False):
  if not gmail_user or not gmail_pswd:
    return -1, '{0}Invalid Gmail Username or Password{1}'.format(RED, RESET)
  else:
    email_from = gmail_user
    email_body = ''
    if edd:
      email_body += 'EDD: {0}\n'.format(edd)
    if state:
      email_body += 'Status: {0}\n'.format(state.title())
    if send_ws:
      email_body += 'Window Sticker Released!\n'

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
      file_name = '{0}.pdf'.format(vin)
      with open(file_name, 'rb') as in_file:
        attachment = MIMEApplication(in_file.read(), Name=file_name)
      attachment['Content-Disposition'] = 'attachment; filename="{0}"'.format(file_name)
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
    except:
      return -1, '{0}FAIL{1}'.format(RED, RESET)

def check_order(args, order, tid):
  global order_str_list

  for i in range(len(COTUS_URL)):
    url = COTUS_URL[i]
    data = get_data(args, order[0], url=url)
    err, msg = format_order_info(data, args.vehicle_summary, args.send_email, url, args.window_sticker)
    if not err:
      break
  the_lock.acquire()
  order_str_list[tid] = msg
  the_lock.release()

def main():
  global order_str_list

  parser = argparse.ArgumentParser()
  parser.add_argument('-o', '--order-number', type=str, help='order number of the car', dest='order_number')
  parser.add_argument('-d', '--dealer-code', type=str, help='dealer code of the order', dest='dealer_code')
  parser.add_argument('-l', '--last-name', type=str, help='customer\'s last name (not used for now)', dest='last_name', default='xxx')
  parser.add_argument('-v', '--vin', type=str, help='VIN of the car', dest='vin')
  parser.add_argument('-s', '--vehicle-summary', help='show vehicle summary', dest='vehicle_summary', action='store_true')
  parser.add_argument('-f', '--file', type=str, help='file with many many VIN\'s', dest='file')
  parser.add_argument('-e', '--send-email', type=str, help='send email if state changed', dest='send_email')
  parser.add_argument('-w', '--window-sticker', help='obtain the window sticker', dest='window_sticker', action='store_true')
  args = parser.parse_args()

  if args.file:
    if not os.path.isfile(args.file):
      print('Invalid VIN file.')
      exit(1)
    else:
      orders = get_orders(args.file)
      threads = []
      for o in orders:
        if o[0] == 'vin':
          if len(o) == 2:
            args.vin = o[1]
          elif len(o) == 3:
            args.vin = o[1]
            args.send_email = o[2]
          else:
            print('Invalid Order.')
            continue
        elif o[0] == 'num':
          if len(o) == 3:
            args.order_number = o[1]
            args.dealer_code = o[2]
          elif len(o) == 4:
            args.order_number = o[1]
            args.dealer_code = o[2]
            args.send_email = o[3]
          else:
            print('Invalid Order.')
            continue
        else:
          print('Invalid Order.')
          continue
        tid = len(threads)
        threads.append(threading.Thread(target=check_order, args=(copy.deepcopy(args), o, tid)))
        order_str_list.append('')
      for t in threads:
        t.start()
      for t in threads:
        t.join()
      for s in order_str_list:
        print(s)
  else:
    for i in range(len(COTUS_URL)):
      url = COTUS_URL[i]
      if args.vin:
        data = get_data(args, 'vin', url=url)
      elif args.order_number and args.dealer_code and args.last_name:
        data = get_data(args, url=url)
      else:
        print('Invalid input!')
        exit(1)
      err, msg = format_order_info(data, args.vehicle_summary, args.send_email, url, args.window_sticker)
      if not err:
        break
    print(msg)

if __name__ == '__main__':
  main()
