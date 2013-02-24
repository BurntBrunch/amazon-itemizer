#!/usr/bin/env python

import imaplib
import pickle
import io
import csv
import re

USER = 'user@gmail.com'
PASS = 'your-gmail-application-password'

# Add option to pull from pickle first


def main():
    # Grab mail messages from Gmail
    orders = []
    messages = get_mail_messages(USER, PASS)

    if messages:
        orders = process_messages(messages)
    if orders:
        print "Exporting orders to CSV file..."
        save_to_csv(orders)
        print "Export complete."


def get_mail_messages(user, pwd):
    "Retrieves mail sent by ship-confirm@amazon.com"
    messages = []

    print 'Opening mail connection...'
    mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
    mail.login(USER, PASS)
    mail.select('Amazon')

    print 'Searching for mail...'
    status, data = mail.search(None, 'FROM', '"ship-confirm@amazon.com"')

    print 'Fetching. This may take a few minutes...'
    for num in data[0].split():
        status, data = mail.fetch(num, '(RFC822)')
        # print 'Message %s\n%s\n' % (num, data[0][1])
        messages.append(data[0][1])

    print 'Closing mail connection...'
    mail.close()
    mail.logout()

    return messages


def process_messages(messages):
    "Parse shipping confirmation emails from Amazon"
    # Procesed messages are stored in orders
    orders = []

    print 'Parsing thru %s messages...' % len(messages)
    for mail in messages:
        line = []

        # Clean up this part
        order_date = re.search(r"(\d+\ \w+\ \d{4})\ \d{2}", mail)
        order_total = re.search(r"\sTotal:\s+\$(\d+.\d+)", mail)
        order_number = re.search(r"(\d+-\d+-\d+)", mail)

        item_type1 = re.findall(r" {3}(\w.+)\n {3}\$(\d+.\d+)", mail)
        item_type2 = re.findall(r"\d(.+)\$(\d+.\d+)", mail)
        # m = re.findall(r"\w[^:]+\s+\$\d+.\d+$", mail)

        if order_date:
            line.append(order_date.group(1))
        else:
            line.append('no date')

        if order_number:
            line.append(order_number.group(1))
        else:
            line.append('no order no')

        if order_total:
            line.append(order_total.group(1))
        else:
            line.append('no total')

        if item_type1:
            for lineitem in item_type1:
                for col in lineitem:
                    col = re.sub('\s+\$\d+.\d+\s+\d', '', col)
                    col = col.replace(' \r', '').replace(', ', '').strip()
                    line.append(col)
        elif item_type2:
            for lineitem in item_type2:
                for col in lineitem:
                    col = re.sub('\s+\$\d+.\d+\s+\d', '', col)
                    col = col.replace(' \r', '').replace(', ', '').strip()
                    line.append(col)
        else:
            #line.append(0)
            line.append('could not find line items')

        # Build the line item
        orders.append(line)

    return orders


def save_to_csv(orders):
    "Saves processed orders into a CSV file"
    with open('amazon_orders.csv', 'wb') as csvfile:
        wr = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL)
        for row in orders:
            wr.writerow(row)


def pickle_messages(messages):
    "Pickle all unprocessed mail messages"
    pickle.dump(messages, open("amazon_messages.p", "wb"))


def unpickle_messages():
    "Rehydrate pickled messages."
    messages = pickle.load(open("amazon_messages.p", "rb"))
    return messages

if __name__ == "__main__":
    main()
