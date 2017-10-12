#!/usr/bin/env python3
# kate: space-indent on; tab-indent off;

""" @package docstring
IMAP Util

Utility class for accessing IMAP server

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

class ImapUtil:

    NAME = 'imaputil'
    VERSION = '0.4'

    def listMailboxes(self, conn):
        """
            @param conn: Active IMAP connection
            @return Returns a list of dict{ 'flags', 'delimiter', 'mailbox') }
        """
        (res, data) = conn.list()
        if res != 'OK':
            raise RuntimeError('Invalid reply: ' + res)
        list_re = re.compile(rb'\((?P<flags>.*)\)\s+"(?P<delimiter>.*)"\s+"?(?P<name>[^"]*)"?')
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

    def listMessages(self, conn):
        """
            List all messages in the given conn and current mailbox.

            @returns a list of message imap identifiers
        """
        (res, data) = conn.search(None, 'ALL')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        msgids = data[0].split()
        return msgids

    def getMessageId(self, conn, imapid):
        """
            returns "Message-ID"
        """
        (res, data) = conn.fetch(imapid, '(BODY.PEEK[HEADER])')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        headers = email.message_from_bytes(data[0][1])
        return headers['Message-ID']

    def getMessage(self, conn, imapid):
        """
            returns full RFC822 message
        """
        (res, data) = conn.fetch(imapid, '(RFC822)')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        return data[0][1]

    def getHeaders(self, conn, imapid):
        """
            Returns message headers
        """
        (res, data) = conn.fetch(imapid, '(BODY[HEADER])')
        if res != 'OK':
            raise RuntimeError('Unvalid reply: ' + res)
        parser = email.parser.HeaderParser()
        return parser.parsestr(data[0][1])

    def getServerType(self, conn):
        """ Try to guess IMAP server type
        @return tuple (type, descr) Type is one of: unknown, exchange, dovecot
        """
        regs = {
            'exchange': re.compile(b'^.*Microsoft Exchange.*$', re.I),
            'dovecot': re.compile(b'^.*(imapfront|dovecot).*$', re.I),
            'courier': re.compile(b'^.*Courier.*$', re.I),
        }
        descr = {
            'exchange': 'MS Exchange',
            'dovecot': 'Dovecot',
            'courier': 'Courier',
        }
        for r in regs.keys():
            if regs[r].match(conn.welcome):
                return ( r, descr[r] )
        return ( 'unknown', 'Unknown ({})'.format(conn.welcome.decode()) )

    def translateFolderName(self, name, srcformat, dstformat):
        """ Translates forlder name from src server format do dst server format """

        # 1. Transpose into dovecot format (use DOT as folder separator), no INBOX. prefix
        if srcformat == 'exchange':
            name = name.replace(b'.', b' ').replace(b'/', b'.')
        elif srcformat == 'courier':
            name = re.sub(b'^INBOX.', '', name, 1)
        elif srcformat == 'dovecot':
            pass
        else:
            pass

        # 2. Transpose into output format
        if dstformat == 'exchange':
            name = name.replace(b'/', b' ').replace(b'.', b'/')
        elif dstformat == 'courier':
            name = b'INBOX.' + name
        elif dstformat == 'dovecot':
            pass
        else:
            pass

        return name
