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

class main:
    
    NAME = 'imapcp'
    VERSION = '0.1'
    
    def run(self):
        
        # Init pretty printer
        pp = pprint.PrettyPrinter(indent = 2)
        
        # Read command line
        usage = "%prog <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>"
        parser = OptionParser(usage=usage, version=self.NAME + ' ' + self.VERSION)

        (options, args) = parser.parse_args()
        
        # Parse mandatory arguments
        if len(args) < 2:
            parser.error("unvalid number of arguments")
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
        
        # Made connections and authenticate
        srcconn = imaplib.IMAP4(src['host'], src['port'])
        srcconn.login(src['user'], src['pass'])
        dstconn = imaplib.IMAP4(dst['host'], dst['port'])
        dstconn.login(dst['user'], dst['pass'])

        print "Source mailboxes:"
        srcfolders = self.__listMailboxes(srcconn)
        pp.pprint(srcfolders)
        
        print "Destination mailboxes:"
        dstfolders = self.__listMailboxes(dstconn)
        pp.pprint(dstfolders)

        # Syncing every source folder
        for f in srcfolders:
            print "Syncing", f['mailbox']
            
            # Create dst mailbox when missing
            dstconn.create(f['mailbox'])
            
            # Select source mailbox readonly
            srcconn.select(f['mailbox'], True)
            dstconn.select(f['mailbox'], False)
            
            # Fetch all destination messages imap IDS
            dstids = self.__listMessages(dstconn)
            print "Found", len(dstids), "messages in destination folder"
            
            # Fetch destination messages ID
            print "Acquiring message IDs..."
            dstmexids = []
            for did in dstids:
                dstmexids.append(self.__getMessageId(dstconn, did))
            print len(dstmexids), "message IDs acquired."
            
            # Fetch all source messages imap IDS
            srcids = self.__listMessages(srcconn)
            print "Found", len(srcids), "messages in source folder"
            
            # Sync data
            for sid in srcids:
                # Get message id
                mid = self.__getMessageId(srcconn, sid)
                if not mid in dstmexids:
                    # Message not found, syncing it
                    print "Copying message", mid
                    mex = self.__getMessage(srcconn, sid)
                    dstconn.append(f['mailbox'], None, None, mex)
                else:
                    print "Skipping message", mid

        # Logout
        srcconn.logout()
        dstconn.logout()

    def __listMailboxes(self, conn):
        """
            @param conn: Active IMAP connection
            @return Returns a list of dict{ 'flags', 'delimiter', 'mailbox') }
        """
        (res, data) = conn.list()
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        list_re = re.compile(r'\((?P<flags>.*)\)\s+"(?P<delimiter>.*)"\s+"(?P<name>.*)"')
        folders = []
        for d in data:
            m = list_re.match(d)
            if not m:
                raise RuntimeError('No match: ' + d)
            flags, delimiter, mailbox = m.groups()
            folders.append({
                'flags': flags,
                'delimiter': delimiter,
                'mailbox': mailbox,
            })
        return folders

    def __listMessages(self, conn):
        """
            List all messages in the given conn and current mailbox.
            
            @returns a list of message imap identifiers
        """
        (res, data) = conn.search(None, 'ALL')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        msgids = data[0].split()
        return msgids

    def __getMessageId(self, conn, imapid):
        """
            returns "Message-ID"
        """
        (res, data) = conn.fetch(imapid, '(BODY.PEEK[HEADER])')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        headers = email.message_from_string(data[0][1])
        return headers['Message-ID']

    def __getMessage(self, conn, imapid):
        """
            returns full RFC822 message
        """
        (res, data) = conn.fetch(imapid, '(RFC822)')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        return data[0][1]

if __name__ == '__main__':
    app = main()
    app.run()
    sys.exit(0)
