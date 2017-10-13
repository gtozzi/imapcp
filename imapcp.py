#!/usr/bin/env python3
# kate: space-indent on; tab-indent off;

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
import datetime

from optparse import OptionParser
from imaputil import ImapUtil

class main(ImapUtil):

    NAME = 'imapcp'
    VERSION = '0.5'

    def run(self):

        # Init pretty printer
        pp = pprint.PrettyPrinter(indent = 2)

        # Read command line
        usage = "%prog <user>:<password>:<host>:<port> <user>:<password>:<host>:<port>"
        parser = OptionParser(usage=usage, version=self.NAME + ' ' + self.VERSION)
        parser.add_option("-e", "--exclude", dest="exclude", action='append',
            help="Exclude folders matching pattern (can be specified multiple times)")
        parser.add_option("-f", "--folder", dest="folder",
            help="Only copy a single folder (use from:to to specify a different destinatin name)")
        parser.add_option("-s", "--simulate", dest="simulate", action='store_true',
            help="Do not perform any task")
        parser.add_option("-k", "--skel", dest="skel", action='store_true',
            help="Only copy folder structure")
        parser.add_option("--from", dest="fr",
            help="Only copy messages older than this date (inclusive)")
        parser.add_option("--to", dest="to",
            help="Only copy messages newer than this date (inclusive)")

        (options, args) = parser.parse_args()

        # Parse exclude list
        excludes = []
        if options.exclude:
            for e in options.exclude:
                excludes.append(re.compile(e.encode()))

        # Parse from/to dates
        fr = None
        if options.fr:
            fr = datetime.date(*[int(i) for i in options.fr.split('-')])
        if fr:
            print("Only copying messages newer than %s (included)" % fr)
        to = None
        if options.to:
            to = datetime.date(*[int(i) for i in options.to.split('-')])
        if to:
            print("Only copying messages older than %s (included)" % to)

        # Parse single folder
        folder = options.folder.split(':') if options.folder else None
        if folder and len(folder) < 2:
            folder.append(folder[0])
        if folder:
            print("Only copying folder %s to folder %s" % (folder[0], folder[1]))

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
        if src['port'] == 993:
            srcconn = imaplib.IMAP4_SSL(src['host'], src['port'])
        else:
            srcconn = imaplib.IMAP4(src['host'], src['port'])
        srcconn.login(src['user'], src['pass'])
        srctype, srcdescr = self.getServerType(srcconn)
        print("Source server type is", srcdescr)

        if dst['port'] == 993:
            dstconn = imaplib.IMAP4_SSL(dst['host'], dst['port'])
        else:
            dstconn = imaplib.IMAP4(dst['host'], dst['port'])
        dstconn.login(dst['user'], dst['pass'])
        dsttype, dstdescr = self.getServerType(dstconn)
        print("Destination server type is", dstdescr)

        print("Source folders:")
        srcfolders = self.listMailboxes(srcconn)
        for f in srcfolders:
             print(f)

        print("Destination folders:")
        dstfolders = self.listMailboxes(dstconn)
        for f in dstfolders:
            print(f)

        # Syncing every source folder
        for f in srcfolders:

            # Translate folder name
            srcfolder = f.name
            dstfolder = f.getPathBytes(dsttype)

            # Check for folder in exclusion/inclusion list
            skip = False
            if folder:
                if folder[0] != srcfolder:
                    skip = True
                elif folder[1]:
                    dstfolder = folder[1]
            else:
                for e in excludes:
                    if e.match(srcfolder):
                        skip = True
                        break
            if skip:
                print("Skipping", srcfolder, "(excluded)")
                continue

            print("Syncing", srcfolder, 'into', dstfolder)

            # Create dst mailbox when missing
            dstconn.create(dstfolder)

            # Select source mailbox readonly
            (res, data) = srcconn.select(srcfolder, True)
            if res == 'NO' and srctype == 'exchange' and 'special mailbox' in data[0]:
                print("Skipping special Microsoft Exchange Mailbox", srcfolder)
                continue
            dstconn.select(dstfolder, False)

            # Stop here if only copying skeleton
            if options.skel:
                print("Skipping message copy")
                continue

            # Fetch all destination messages imap IDS
            dstids = self.listMessages(dstconn)
            print("Found", len(dstids), "messages in destination folder")

            # Fetch destination messages ID
            print("Acquiring destination message IDs...", end='', flush=True)
            dstmexids = []
            for idx, did in enumerate(dstids):
                if idx % 100 == 0:
                    print('.', end='', flush=True)
                dstmexids.append(self.getMessageId(dstconn, did))
            print(len(dstmexids), "message IDs acquired.")

            # Fetch all source messages imap IDS
            srcids = self.listMessages(srcconn)
            print("Found", len(srcids), "messages in source folder")

            # Sync data
            for sid in srcids:
                # Check for date filter
                if fr or to:
                    h = self.getHeaders(srcconn, sid)
                    if 'date' not in h:
                        continue
                    d = email.utils.parsedate(h['date'])
                    if not d:
                        continue
                    date = datetime.date(d[0], d[1], d[2])
                    if fr and date < fr:
                        continue
                    if to and date > to:
                        continue
                # Get message id
                mid = self.getMessageId(srcconn, sid)
                if not mid in dstmexids:
                    # Message not found, syncing it
                    print("Copying message", mid)
                    if not options.simulate:
                        mex = self.getMessage(srcconn, sid)
                        dstconn.append(dstfolder, None, None, mex)
                else:
                    print("Skipping message", mid)

        # Logout
        srcconn.logout()
        dstconn.logout()

        if options.simulate:
            print("Simulated run, no action taken")


if __name__ == '__main__':
    app = main()
    app.run()
    sys.exit(0)
