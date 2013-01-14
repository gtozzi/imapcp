#!/usr/bin/env python
# -*- coding: utf-8 -*-

""" @package docstring
IMAP Copy

Copy emails and folders from an IMAP account to another.
Creates missing folders and skips existing messages (using message-id).

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

from optparse import OptionParser
from imaputil import ImapUtil

class main(ImapUtil):
    
    NAME = 'imapcp'
    VERSION = '0.2'
    
    def run(self):
        
        # Init pretty printer
        pp = pprint.PrettyPrinter(indent = 2)
        
        # Read command line
        usage = "%prog <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>"
        parser = OptionParser(usage=usage, version=self.NAME + ' ' + self.VERSION)
        parser.add_option("-e", "--exclude", dest="exclude", action='append',
            help="Exclude folders matching pattern (can be specified multiple times)")

        (options, args) = parser.parse_args()
        
        # Parse exclude list
        excludes = []
        if options.exclude:
            for e in options.exclude:
                excludes.append(re.compile(e))
        
        # Parse mandatory arguments
        if len(args) < 2:
            parser.error("invalid number of arguments")
        src = args[0].split(':')
        src = {
            'user': src[0],
            'pass': src[1],
            'host': src[2] if len(src) > 2 else 'localhost',
            'port': int(src[3]) if len(src) > 3 else 143,
        }
        dst = args[1].split(':')
        dst = {
            'user': dst[0],
            'pass': dst[1],
            'host': dst[2] if len(dst) > 2 else 'localhost',
            'port': int(dst[3]) if len(dst) > 3 else 143,
        }
        
        # Make connections and authenticate
        srcconn = imaplib.IMAP4(src['host'], src['port'])
        srcconn.login(src['user'], src['pass'])
        srctype = self.getServerType(srcconn)
        print "Source server type is", srctype
        
        dstconn = imaplib.IMAP4(dst['host'], dst['port'])
        dstconn.login(dst['user'], dst['pass'])
        dsttype = self.getServerType(dstconn)
        print "Destination server type is", dsttype

        print "Source mailboxes:"
        srcfolders = self.listMailboxes(srcconn)
        pp.pprint(srcfolders)
        
        print "Destination mailboxes:"
        dstfolders = self.listMailboxes(dstconn)
        pp.pprint(dstfolders)

        # Syncing every source folder
        for f in srcfolders:
            
            # Translate folder name
            srcfolder = f['mailbox']
            dstfolder = self.translateFolderName(f['mailbox'], srctype, dsttype)
            
            # Check for folder in exclusion list
            skip = False
            for e in excludes:
                if e.match(srcfolder):
                    skip = True
                    break
            if skip:
                print "Skipping", srcfolder, "(excluded)"
                continue
            
            print "Syncing", srcfolder, 'into', dstfolder
            
            # Create dst mailbox when missing
            dstconn.create(dstfolder)
            
            # Select source mailbox readonly
            (res, data) = srcconn.select(srcfolder, True)
            if res == 'NO' and srctype == 'exchange' and 'special mailbox' in data[0]:
                print "Skipping special Microsoft Exchange Mailbox", srcfolder
                continue
            dstconn.select(dstfolder, False)
            
            # Fetch all destination messages imap IDS
            dstids = self.listMessages(dstconn)
            print "Found", len(dstids), "messages in destination folder"
            
            # Fetch destination messages ID
            print "Acquiring message IDs..."
            dstmexids = []
            for did in dstids:
                dstmexids.append(self.getMessageId(dstconn, did))
            print len(dstmexids), "message IDs acquired."
            
            # Fetch all source messages imap IDS
            srcids = self.listMessages(srcconn)
            print "Found", len(srcids), "messages in source folder"
            
            # Sync data
            for sid in srcids:
                # Get message id
                mid = self.getMessageId(srcconn, sid)
                if not mid in dstmexids:
                    # Message not found, syncing it
                    print "Copying message", mid
                    mex = self.getMessage(srcconn, sid)
                    dstconn.append(dstfolder, None, None, mex)
                else:
                    print "Skipping message", mid

        # Logout
        srcconn.logout()
        dstconn.logout()

if __name__ == '__main__':
    app = main()
    app.run()
    sys.exit(0)
