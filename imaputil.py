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


class MailFolder:
    """ A Mail Folder representation """

    def __init__(self, srvtype, flags, delimiter, name):
        self.srvtype = srvtype
        self.delimiter = delimiter
        self.flags = flags
        self.name = name

    def getPath(self):
        """
            @return tuple: standardized path as a tuple
        """
        path = self.name.split(self.delimiter)
        if self.srvtype == ImapUtil.TYPE_COURIER:
            # Remove trailing inbox
            if path[0] != b'INBOX':
                raise ValueError('Courier path must start with inbox: {}'.format(path[0]))
            if len(path) > 1:
                path = path[1:]
        return tuple(path)

    def getPathBytes(self, srvtype=None, trim=False):
        """
            @return path translated for given server type
        """
        if srvtype is None:
            srvtype = self.srvtype

        path = self.getPath()
        if trim:
            path = map(lambda i: i.strip(), path)

        if srvtype == ImapUtil.TYPE_EXCHANGE:
            # Use slash
            return b'/'.join(path)
        else:
            # Use dot
            path = b'.'.join(path)
            if srvtype == ImapUtil.TYPE_COURIER and path != b'INBOX':
                # Append INBOX.
                path = b'INBOX.' + path
            return path

    def __bytes__(self):
        return self.name

    def __repr__(self):
        return str(self.getPath())


class ImapUtil:

    NAME = 'imaputil'
    VERSION = '0.4'

    TYPE_EXCHANGE = 'exchange'
    TYPE_DOVECOT = 'dovecot'
    TYPE_COURIER = 'courier'
    TYPE_UNKNOWN = 'unknown'

    def listMailboxes(self, conn):
        """
            @param conn: Active IMAP connection
            @return Returns a list of Mailbox objects
        """
        srvtype, srvdescr = self.getServerType(conn)
        (res, data) = conn.list()
        if res != 'OK':
            raise RuntimeError('Invalid reply: ' + res)
        list_re = re.compile(rb'\((?P<flags>.*)\)\s+"(?P<delimiter>.*)"\s+"?(?P<name>[^"]*)"?')
        folders = []
        for d in data:
            m = list_re.match(d)
            if not m:
                raise RuntimeError('No match: ' + d)
            flags, delimiter, name = m.groups()
            folders.append(MailFolder(srvtype, flags, delimiter, name))
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
            self.TYPE_EXCHANGE: re.compile(b'^.*Microsoft Exchange.*$', re.I),
            self.TYPE_DOVECOT: re.compile(b'^.*(imapfront|dovecot).*$', re.I),
            self.TYPE_COURIER: re.compile(b'^.*Courier.*$', re.I),
        }
        descr = {
            self.TYPE_EXCHANGE: 'MS Exchange',
            self.TYPE_DOVECOT: 'Dovecot',
            self.TYPE_COURIER: 'Courier',
        }
        for r in regs.keys():
            if regs[r].match(conn.welcome):
                return ( r, descr[r] )
        return ( self.TYPE_UNKNOWN, 'Unknown ({})'.format(conn.welcome.decode()) )

    def translateFolderName(self, folder, srcformat, dstformat):
        """ Translates folder name from src server format do dst server format """

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
