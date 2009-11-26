"""
OWA MAIL CHECK ver. 1.0 modified by riccardo.donato@gmail.com based on:
iPhone-pop3-outlook
iPhone POP e-mail server for Microsoft Outlook Web Access scraper

This wraps the Outlook Web Access scraper by providing a POP interface to it.

Run this file from the command line to start the server on localhost port 110.

OWA web address is now passed to the script in the username rather than hard-coded into
the script. Login as follows:

Host Name: localhost
User Name: https://mail.yourcompany.com/exchange yourdomain\yourusername
Password: yourpassword

Note the single space between your OWA web address and your username. "yourdomain\" may not be
required depending on your company's OWA configuration.
"""

# Change History:
# 0.0.1 - Adrian Holvaty original script
# 0.0.1.1 - 2007-08-20 - lh <lh@mail.saabnet.com> disabled delete functionality
# 0.0.2 - 2007-08-23 - lh <lh@mail.saabnet.com> added debug code, disabled quit_after_one,
#         replaced 'welcome' message to fix script on iPhone, other minor tweaks
# 0.0.3 - 2007-08-25 - lh <lh@mail.saabnet.com> added TOP and UIDL support - tested working on iPhone
# 0.0.4 - 2007-08-31 - lh <lh@mail.saabnet.com> added text-based local email storage to speed up script,
#                      IP-based security, removed extraneous code
# 0.1.0 - 2007-09-26 - lh <lh@mail.saabnet.com> Working FORMS-based version
# 
# Based on gmailpopd.py by follower@myrealbox.com,
# which was in turn based on smtpd.py by Barry Warsaw.
#
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

import asyncore, asynchat, socket, sys, re, os

# relative import
from owa_scraper import InvalidLogin, OutlookWebScraper

__version__ = 'Python FORMS Outlook Web Access POP3 proxy version 0.1.0'

TERMINATOR = '\r\n'
#WEBMAIL_SERVER = ''

def quote_dots(lines):
    for line in lines:
        if line.startswith("."):
            line = "." + line
        yield line

class POPChannel(asynchat.async_chat):
    #def __init__(self, conn, quit_after_one):
    def __init__(self, conn):
        #self.quit_after_one = quit_after_one
        asynchat.async_chat.__init__(self, conn)
        self.__line = []
        self.push('+OK Outlook Web Access POP3 Proxy')
        #self.push('+OK %s %s' % (socket.getfqdn(), __version__))
        print "Sent welcome message"
        self.set_terminator(TERMINATOR)
        self._activeDataChannel = None

    # Overrides base class for convenience
    def push(self, msg):
        asynchat.async_chat.push(self, msg + TERMINATOR)

    # Implementation of base class abstract method
    def collect_incoming_data(self, data):
        self.__line.append(data)

    # Implementation of base class abstract method
    def found_terminator(self):
        line = ''.join(self.__line)
        self.__line = []
        if not line:
            self.push('500 Error: bad syntax')
            print "error 1"
            return
        method = None
        i = line.find(' ')
        if i < 0:
            command = line.upper()
            arg = None
        else:
            command = line[:i].upper()
            arg = line[i+1:].strip()
        method = getattr(self, 'pop_' + command, None)
        if not method:
            self.push('-ERR Error : command "%s" not implemented' % command)
            print "error 2: %s" % command
            return
        method(arg)
        return

    def pop_USER(self, user):
        # Logs in any username.
        if not user:
            self.push('-ERR: Syntax: USER username')
            print "error 3"
        else:
            login = user.split(' ')
            print login
            self.webmail_server = login[0]
            self.username = login[1] # Store for later.
            self.push('+OK Password required')
            print "User Name sent, asking for password"

    def pop_PASS(self, password=''):
        self.scraper = OutlookWebScraper(self.webmail_server, self.username, password)
        print self.webmail_server
        print self.username
        #print password
        
        try:
            self.scraper.login()
        except InvalidLogin:
            self.push('-ERR Login failed. (Wrong username/password?)')
            print "error 4"
        else:
            self.push('+OK User logged in')
            print "User logged in"
            
            print "Accessing inbox"        
            self.inbox_cache = self.scraper.inbox()

            #load message list index from last session
            print "Checking outlook webmail for new mails.."
            msg = self.scraper.inbox()
            #DEBUG print msg

    def pop_STAT(self, arg):
        # dropbox_size = sum([len(msg) for msg in self.msg_cache])
        msg = self.scraper.inbox()
        if msg == "YES" : 
           self.push('+OK 1 320\r\n')         
        else :
           self.push('+OK 0 0\r\n')
        print "Sent STAT"

    def pop_LIST(self, arg):       
        msg = self.scraper.inbox()
        if msg == "YES" :
         self.push('+OK 1 message (320 octets)')
         self.push('1 320')
         self.push('.')
        else :
         self.push('+OK --')
         self.push('.')
        print "Sent LIST"

    def pop_RETR(self, arg):
        print "DEBUG"
        if not arg:
            self.push('-ERR: Syntax: RETR msg')
            print "error 5"

        else:
            # TODO: Check request is in range.
            #msg_index = int(arg) - 1
            msg = self.inbox_cache
            self.push('+OK ')
            self.push('320 octets')
            self.push('From: owacheck@owa_mailcheck')
            self.push('Subject: You have unreaded owa-mails!')          
            self.push('Content-Type: text/plain; charset=\"us-ascii\"')
            self.push('Content-Transfer-Encoding: quoted-printable')
            self.push('MIME-Version: 1.0')
            self.push('Content-Disposition: inline')
            self.push('.')
            print "sent message"

            # Delete the message
            #self.scraper.delete_message(msg_id)

    def pop_QUIT(self, arg):
        #print "DEBUG"
        self.push('+OK Goodbye')
        print "User quit"
        self.close_when_done()
        #if self.quit_after_one:
            # This SystemExit gets propogated to handle_error(),
            # which stops the program. Slightly hackish.
            #raise SystemExit
            
    def pop_UIDL(self, arg):
    	#TODO clean up this code
        #print "DEBUG"
        if not arg:
            print "got UIDL request"
            num_messages = 1
            self.push('+OK')
            #search each message for UID
            self.push('1 ABC')                
            self.push('.')
            print "Sent UIDL"
        else:
            # TODO: Handle per-msg LIST commands
            raise NotImplementedError
            
    def pop_DELE(self, arg):
           print "got DELE request"
           self.push('+OK message 1 deleted')
           print "sent DELE response"
    

    def pop_TOP(self, arg):
    	#TODO clean up this code
        print "got TOP request %s" % arg
        argt = arg.split(' ')
        #message ID (internal)
        mid = int(argt[0])
        #number of lines requested (only 0 & 40 supported)
        lin = int(argt[1])
        
        if not arg:
            self.push('-ERR: Syntax: RETR msg')
            print "error 5"
        else:   
          self.push('+OK')               
          self.push('.')
        print "sent TOP"

    def handle_error(self):
        #if self.quit_after_one:
            #sys.exit(0) # Exit.
        #else:
        asynchat.async_chat.handle_error(self)

class POP3Proxy(asyncore.dispatcher):
    #def __init__(self, localaddr, quit_after_one):
    def __init__(self, ip, port):
        """
        localaddr is a tuple of (ip_address, port).

        quit_after_one is a boolean specifying whether the server should quit
        after serving one session.
        """
        self.ip= ip
        self.port = port
        
        #self.quit_after_one = quit_after_one
        asyncore.dispatcher.__init__(self)
        self.create_socket(socket.AF_INET, socket.SOCK_STREAM)
        # try to re-use a server port if possible
        self.set_reuse_addr()
        self.bind((ip, port))
        self.listen(5)

    def handle_accept(self):
        conn, addr = self.accept()
        print addr
        if addr[0] == '127.0.0.1':
        	#channel = POPChannel(conn, self.quit_after_one)
        	channel = POPChannel(conn)
        else:
        	print "Outside IP - rejected!"
        	
def createDaemon():
	'''Funzione che crea un demone per eseguire un determinato programma...'''
	#copied from http://snippets.dzone.com/posts/show/1532
	import os
	
	# create - fork 1
	try:
		if os.fork() > 0: os._exit(0) # exit father...
	except OSError, error:
		print 'fork #1 failed: %d (%s)' % (error.errno, error.strerror)
		os._exit(1)

	# it separates the son from the father
	os.chdir('/')
	os.setsid()
	os.umask(0)

	# create - fork 2
	try:
		pid = os.fork()
		if pid > 0:
			print 'Daemon PID %d' % pid
			os._exit(0)
	except OSError, error:
		print 'fork #2 failed: %d (%s)' % (error.errno, error.strerror)
		os._exit(1)

	asyncore.loop() # function demo

if __name__ == '__main__':
    from optparse import OptionParser
    parser = OptionParser("usage: %prog [options]")
    parser.add_option('--once', action='store_true', dest='once',
        help='Serve one POP transaction and then quit. (Server runs forever by default.)')
    options, args = parser.parse_args()
    #proxy = POP3Proxy(('', 10110), options.once is True)
    proxy = POP3Proxy('', 110)
    print "POP Server running"
    try:
        asyncore.loop()
        #createDaemon()
    except KeyboardInterrupt:
        pass
