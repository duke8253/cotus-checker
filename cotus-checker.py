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

from grab import Grab
import re
import argparse
import os

RED    = '\033[1;31m'
GREEN  = '\033[1;32m'
YELLOW = '\033[1;33m'
BLUE   = '\033[1;34m'
PURPLE = '\033[1;35m'
CYAN   = '\033[1;36m'
WHITE  = '\033[1;37m'
RESET  = '\033[0;0m'

COTUS_URL1 = 'http://wwwqa.cotus.ford.com'
COTUS_URL2 = 'http://www.cotus.ford.com'

order_states = {
  0: 'In Order Processing:',
  1: 'In Production:',
  2: 'Awaiting Shipment:',
  3: 'In Transit:',
  4: 'Delivered:'
}

def get_vins(file_name):
  with open(file_name, 'r') as in_file:
    lines = in_file.readlines()
  vins = []
  for l in lines:
    vins.append(l.replace('\n', '').strip())
  return vins

def setup_order_num(args, g):
  g.doc.set_input('orderTrackingInputType', 'orderNumberInput')
  g.doc.set_input('orderNumber', args.order_number)
  g.doc.set_input('dealerCode', args.dealer_code)
  g.doc.set_input('customerLastName', args.last_name)

def setup_vin(args, g):
  g.doc.set_input('orderTrackingInputType', 'vin')
  g.doc.set_input('vin', args.vin)

def get_data(args, which_one='', url=COTUS_URL1):
  g = Grab()
  g.go(url)
  if which_one == 'vin':
    setup_vin(args, g)
  else:
    setup_order_num(args, g)
  return g.doc.submit().unicode_body().replace('\n', '').replace('\r', '')

def get_order_info(data):
  order_info = {
    'vehicle_name': re.search(u'class="vehicleName">(.*?)</span>', data).group(1).strip(),
    'order_date': re.search(u'class="orderDate">(.*?)</span>', data).group(1).strip(),
    'order_num': re.search(u'class="orderNumber">(.*?)</span>', data).group(1).strip(),
    'order_vin': re.search(u'class="vin">(.*?)</span>', data).group(1).strip(),
    'order_edd': re.search(u'id="hidden-estimated-delivery-date" data-part="(.*?)"', data).group(1).strip(),
    'current_state': re.search(u'"selectedStepName":(.*?)"surveyOn"', data).group(1).replace(',', '').replace('"', '').strip()
  }

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

def format_order_info(data, vehicle_summary=False):
  try:
    error_msg = re.search(u'class="top-level-error enabled">(.*?)</p>', data).group(1).strip()
    return 1, error_msg

  except AttributeError:
    order_info = get_order_info(data)

    order_str = 'Order Information:\n'
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Vehicle Name:', GREEN, order_info['vehicle_name'], RESET)
    order_str += '  {0: <21}{1}\n'.format('Ordered On:', order_info['order_date'])
    order_str += '  {0: <21}{1}\n'.format('Order Number:', order_info['order_num'])
    order_str += '  {0: <21}{1}\n'.format('VIN:', order_info['order_vin'])
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Estimated Delivery:', CYAN, 'N/A' if not order_info['order_edd'] else order_info['order_edd'], RESET)
    order_str += '  {0: <21}{1}{2}{3}\n'.format('Current State:', RED, order_info['current_state'], RESET)

    for i in range(5):
      try:
        order_str += '  {0: <21}{1}\n'.format(order_states[i], order_info['state_dates'][i])
      except IndexError:
        order_str += '  {0: <21}{1}\n'.format(order_states[i], 'N/A')

    if vehicle_summary:
      order_str += '\n  Vehicle Summary:\n'
      for each in order_info['vehicle_summary']:
        order_str += '    {0}\n'.format(each)
      order_str += '\n'

    return 0, order_str

def main():
  parser = argparse.ArgumentParser()
  parser.add_argument('-o', '--order-number', type=str, help='order number of the car', dest='order_number')
  parser.add_argument('-d', '--dealer-code', type=str, help='dealer code of the order', dest='dealer_code')
  parser.add_argument('-l', '--last-name', type=str, help='customer\'s last name (not used for now)', dest='last_name', default='xxx')
  parser.add_argument('-v', '--vin', type=str, help='VIN of the car', dest='vin')
  parser.add_argument('-s', '--vehicle-summary', help='show vehicle summary', dest='vehicle_summary', action='store_true')
  parser.add_argument('-f', '--file', type=str, help='file with many many VIN\'s', dest='file')
  args = parser.parse_args()

  if args.file:
    if not os.path.isfile(args.file):
      print('Invalid VIN file.')
      exit(1)
    else:
      vins = get_vins(args.file)
      for v in vins:
        args.vin = v
        data = get_data(args, 'vin', url=COTUS_URL1)
        err, msg = format_order_info(data, args.vehicle_summary)
        if err:
          data = get_data(args, 'vin', url=COTUS_URL2)
          err, msg = format_order_info(data, args.vehicle_summary)
        print(msg)
  else:
    if args.vin:
      data = get_data(args, 'vin', url=COTUS_URL1)
    else:
      data = get_data(args, url=COTUS_URL1)
    err, msg = format_order_info(data, args.vehicle_summary)
    if err:
      if args.vin:
        data = get_data(args, 'vin', url=COTUS_URL2)
      else:
        data = get_data(args, url=COTUS_URL2)
      err, msg = format_order_info(data, args.vehicle_summary)
      print(msg)

if __name__ == '__main__':
  main()
