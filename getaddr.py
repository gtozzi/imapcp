#!/usr/bin/env python
# -*- coding: utf-8 -*-

""" @package docstring
IMAP GetAddr

Gets email addresses from a source IMAP server and saves them in a CSV file

Source IMAP is always accessed READ-ONLY.

@author Gabriele Tozzi <gabriele@tozzi.eu>

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
"""

import imaplib
import sys
import re
import pprint
import email
import csv

from optparse import OptionParser
from imaputil import ImapUtil

class main(ImapUtil):
    
    NAME = 'getaddr'
    VERSION = '0.1'
    
    def run(self):
        
        # Init pretty printer
        pp = pprint.PrettyPrinter(indent = 2)
        
        # Read command line
        usage = "%prog <user>:<password>:<host>:<port>"
        parser = OptionParser(usage=usage, version=self.NAME + ' ' + self.VERSION)

        (options, args) = parser.parse_args()
        
        # Parse mandatory arguments
        if len(args) < 1:
            parser.error("invalid number of arguments")
        src = args[0].split(':')
        src = {
            'user': src[0],
            'pass': src[1],
            'host': src[2] if len(src) > 2 else 'localhost',
            'port': int(src[3]) if len(src) > 3 else 143,
        }
        
        # Make connections and authenticate
        srcconn = imaplib.IMAP4(src['host'], src['port'])
        srcconn.login(src['user'], src['pass'])
        srctype = self.getServerType(srcconn)
        print "Source server type is", srctype

        print "Source mailboxes:"
        srcfolders = self.listMailboxes(srcconn)
        pp.pprint(srcfolders)

        addrs = []
        addrre = re.compile('\<([^>]+@[^>]+)\>')

        # Reading every source folder
        for f in srcfolders:
            
            srcfolder = f['mailbox']
            
            print "Reading from", srcfolder
                        
            # Select source mailbox readonly
            (res, data) = srcconn.select(srcfolder, True)
            if res == 'NO' and srctype == 'exchange' and 'special mailbox' in data[0]:
                print "Skipping special Microsoft Exchange Mailbox", srcfolder
                continue
            
            # Fetch all source messages imap IDS
            srcids = self.listMessages(srcconn)
            print "Found", len(srcids), "messages in source folder"
            
            # Output data
            for sid in srcids:
                mid = self.getMessageId(srcconn, sid)
                print "Reading message", mid
                mex = self.getHeaders(srcconn, sid)
                if mex['From']:
                    f = addrre.search(mex['From'])
                    if f:
                        f = f.group(1).strip()
                else:
                    f = None
                if mex['To']:
                    t = addrre.search(mex['To'])
                    if t:
                        t = t.group(1).strip()
                else:
                    t = None
                print f, t
                if not f in addrs:
                    addrs.append(f)
                if not t in addrs:
                    addrs.append(t)
        
        # Save addresses
        with open('out.csv', 'wb') as csvfile:
            writer = csv.writer(csvfile)
            for addr in addrs:
                writer.writerow([addr, ])
            
        # Logout
        srcconn.logout()

if __name__ == '__main__':
    app = main()
    app.run()
    sys.exit(0)
