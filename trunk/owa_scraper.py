"""
OWA MAIL CHECK ver 1.0 modified by riccardo.donato@gmail.com based on:
Microsoft Outlook Web Access scraper FORMS VERSION

Retrieves full, raw e-mails from Microsoft Outlook Web Access by
screen scraping. Can do the following:

    * Log into a Microsoft Outlook Web Access account with a given username
      and password.
    * Retrieve all e-mail IDs from the first page of your Inbox.
    * Retrieve the full, raw source of the e-mail with a given ID.
    * Delete an e-mail with a given ID (technically, move it to the "Deleted
      Items" folder).

The main class you use is OutlookWebScraper. See the docstrings in the code
and the "sample usage" section below.

This module does no caching. Each time you retrieve something, it does a fresh
HTTP request. It does cache your session, though, so that you only have to log
in once.
"""

# Documentation / sample usage:
#
# # Throws InvalidLogin exception for invalid username/password.
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'invalid password')
# >>> s.login()
# Traceback (most recent call last):
#     ...
# scraper.InvalidLogin
#
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'correct password')
# >>> s.login()
#
# # Display IDs of messages in the inbox.
# >>> s.inbox()
# ['/Inbox/Hey%20there.EML', '/Inbox/test-3.EML']
#
# # Display IDs of messages in the "sent items" folder.
# >>> s.get_folder('sent items')
# ['/Sent%20Items/test-2.EML']
#
# # Display the raw source of a particular message.
# >>> print s.get_message('/Inbox/Hey%20there.EML')
# [...]
#
# # Delete a message.
# >>> s.delete_message('/Inbox/Hey%20there.EML')

# Modified by lh to work with Python 2.5
# September 2007
#
# Based on Outlook Web Access Scraper, version 0.1
# Copyright (C) 2006 Adrian Holovaty <holovaty@gmail.com>
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation; either version 2 of the License, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 59 Temple
# Place, Suite 330, Boston, MA 02111-1307 USA

import re, socket, urlparse, urllib, urllib2, cookielib

__version__ = '0.1.2'
__author__ = 'lh <lh@mail.saabnet.com>'

socket.setdefaulttimeout(15)

class InvalidLogin(Exception):
    pass

class RetrievalError(Exception):
    pass

class CookieScraper(object):
    "Scraper that keeps track of getting and setting cookies."
    def __init__(self):
        #self._cookies = SimpleCookie()
        cj = cookielib.CookieJar()
        self.opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))

    def get_page(self, url, post_data=None, headers=()):
        """
        Helper method that gets the given URL, handling the sending and storing
        of cookies. Returns the requested page as a string.
        """
        
        #req = urllib2.Request(url, post_data)
        #opener.add_header('Cookie', self._cookies.output(attrs=[], header='').strip())

        #create request
        req = urllib2.Request(url, post_data)
        
        for k, v in headers:
            req.add_header(k, v)
        try:
            f = self.opener.open(req)
        except IOError, e:
            print e
            if e[1] == 302:
                # Got a 302 redirect, but check for cookies before redirecting.
                # e[3] is a httplib.HTTPMessage instance.
                if e[3].dict.has_key('set-cookie'):
                    self._cookies.load(e[3].dict['set-cookie'])
                return self.get_page(e[3].getheader('location'))
            else:
                raise
        #if f.headers.dict.has_key('set-cookie'):
        #    self._cookies.load(f.headers.dict['set-cookie'])
        return f.read()

class OutlookWebScraper(CookieScraper):
    def __init__(self, domain, username, password):
        self.domain = domain
        self.username, self.password = username, password
        self.is_logged_in = False
        self.base_href = None
        super(OutlookWebScraper, self).__init__()

    def login(self):
        url = urlparse.urljoin(self.domain, 'exchweb/bin/auth/owaauth.dll')
        html = self.get_page(url, urllib.urlencode({
            'destination': urlparse.urljoin(self.domain, 'exchange/'),
            'flags': '0',
            'username': self.username,
            'password': self.password,
            'SubmitCreds': 'Log On',
            'forcedownlevel': '0',
            'trusted': '4',
        }))
        if 'You could not be logged on to Outlook Web Access' in html:
            raise InvalidLogin
        #m = re.search(r'(?i)<BASE href="([^"]*)">', html)
        self.is_logged_in = True
        # take a look to your html source..it depends on owa implementation (language,etc.)
        #m = re.search("Non letto", html)   
        m = re.search("fld sl bld",html) or re.search("Non letto",html)     
        if not m:
             print 'No new message'
             self.base_href = 'NO'
        else :
        #self.c = self._cookies.output(attrs=[], header='').strip()
             print 'We have a new message'
             self.base_href = 'YES'

    def inbox(self):
        """
        MODIFIED by rdonato
        Returns the message IDs for all messages on the first page of the
        Inbox, regardless of whether they've already been read.
        """
        return self.get_folder('Inbox')

    def get_folder(self, folder_name):
        """
        MODIFIED
        Returns the message IDs for all messages on the first page of the
        folder with the given name, regardless of whether the messages have
        already been read. The folder name is case insensitive.
        """
        if not self.is_logged_in: self.login()
        #url = self.base_href + urllib.quote(folder_name) + '/?Cmd=contents'
        url = self.base_href
        #html = self.get_page(url)
        #message_urls = re.findall(r'(?i)NAME=MsgID value="([^"]*)"', html)
        message_urls = self.base_href
        return message_urls

    def get_message(self, msgid):
        "Returns the inbox new mail response."
        if not self.is_logged_in: self.login()
        return self.base_href
