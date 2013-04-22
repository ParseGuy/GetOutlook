#!/usr/bin/python
# -*- coding: utf-8 -*-

# GetOutlook, by Parse Guy. 2013
# https://github.com/ParseGuy/GetOutlook

from optparse import OptionParser
from configobj import ConfigObj
from cStringIO import StringIO
from email.generator import Generator

import logging
import re
import urllib
import sys
import HTMLParser
import email
import urllib2
import sys
import HTMLParser
import shelve
import codecs
import json

# extra libraries needed
import mechanize

logger = logging.getLogger()


class Outlook:
    ''' Main class of GetOutlook '''
    def __init__(self):
        self.setup()
    description = "Outlook mail fetcher"
    version = "1.01"

    def setup(self):
        ''' Initializes class with Browser object '''
        self.cj = mechanize.LWPCookieJar()
        self.br = mechanize.Browser()
        self.br.set_cookiejar(self.cj)
        self.br.set_handle_equiv(True)
        self.br.set_handle_gzip(False)
        self.br.set_handle_redirect(True)
        self.br.set_handle_referer(True)
        self.br.set_handle_robots(False)
        self.br.set_handle_refresh(
            mechanize._http.HTTPRefreshProcessor(), max_time=1)
        self.br.addheaders = [('User-Agent', "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.8.1.5) Gecko/20061201 Firefox/2.0.0.5 (Ubuntu-feisty)")]

        self.redirurl = 'https://mail.live.com/default.aspx?rru=inbox'

        self.htmlparser = HTMLParser.HTMLParser()

    def parseargs(self):
        ''' Parse arguments of commandline and config file '''
        parser = OptionParser()

        parser.add_option("--config-file", action="store", type="string",
                          dest="configfile", help="Configuration file (mandatory)")
        parser.add_option("--verbosity", action="store", type="int", default=1, dest="verbosity", help="Verbosity of messages (1/2/10/100)")

        (self.options, args) = parser.parse_args()

        if not self.options.configfile:
            parser.error('Missing configuration file')

        # parse the configuration file
        self.configs = ConfigObj(self.options.configfile)

        loglevel = logging.INFO
        # convert old verbosity to new logging
        if self.options.verbosity > 99:
            # include all traffic
            self.br.set_debug_http(True)
            self.br.set_debug_redirects(True)
            self.br.set_debug_responses(True)
            loglevel = logging.DEBUG
        elif self.options.verbosity > 9:
            loglevel = logging.DEBUG
        elif self.options.verbosity > 1:
            loglevel = logging.INFO

        # start the logging with requested level
        logging.basicConfig(level=loglevel, stream=sys.stdout,
                            format='%(levelname)s:%(message)s')

        # check config for options we really need
        configneeded = ['Username', 'Password', 'Domain',
                        'DestinationDir', 'StatusFile']
        for need in configneeded:
            if not need in self.configs:
                logger.error('Missing %s in configuration file' % need)
                return False

        return True

    def getpage(self, url, params=None):
        ''' Get a page from the internet '''
        logger.debug('Fetching: ' + url)
        data = None
        if params:
            data = urllib.urlencode(params)
            logger.debug('Using post params: ' + data)
        error = False
        try:
            response = self.br.open(url, data, timeout=10.0)
        except IOError, e:
            if hasattr(e, 'reason'):
                logger.error("Unable to fetch url '%s': %s", url, e.reason)
            if hasattr(e, 'code'):
                logger.error("url '%s' gives code: %d", url, e.code)
            raise
        return response

    def getcookie(self, name):
        ''' gets cookie from jar '''
        for cookie in self.cj:
            if cookie.name == name:
                return cookie.value
        return None

    def findvar(self, content, name, regex, regexmod=0, logerror=True):
        ''' finds one or more items with regex from an object '''
        m = re.search(regex, content, regexmod)
        if m:
            v = m.groups()
            logger.debug("Found %s:%s", name, v)
            if len(v) == 1:
                return v[0]
            return v
        else:
            if logerror:
                logger.error('Unable to find %s. Regex: %s. Content: %r',
                             name, regex, content)
            raise ValueError('NotFoundVar')

    def login(self):
        ''' Try to login at the Outlook site '''
        # get the start page
        logger.info('Get login page of outlook')
        r = self.getpage('https://mail.live.com').read()

        baseHref = self.findvar(r, 'baseHref', '<base\s+href=\"([^\"]+)\"')

        # retrieve javascript for additional information
        jslist = re.findall(
            '<script\s+type=\"text\/javascript\"\s+src=\"([^\"]+)\"', r)
        if jslist:
            logger.debug('Found javascript hrefs: %s', str(jslist))
        else:
            iogging.error('Unable to find javascript hrefs: %s', r)
            return False

        for js in jslist:
            if not re.match('http', js, re.I):
                js = baseHref + js
                pi
            jspage = self.getpage(js).read()
            r += jspage

        # find information for the login form
        loginurl = self.findvar(
            r, 'loginurl', "(https://login.live.com/ppsecure/post.srf[^']*)'")
        ppsx = self.findvar(r, 'PPSX', "AP:'(P[^']*)", re.S) # h: or F: previous
        ppft = self.findvar(r, 'PPFT', '<\s*input\s+.*name=\"PPFT\"\s+id="\S+"?\s+value=\"(\S*)\"')

        # only use first part of passwd if contains =
        passwd = self.configs['Password'].split('=')[0]
        login = ''.join(
            [self.configs['Username'], '@', self.configs['Domain']])

        # do the post request
        r = self.getpage(loginurl,
                         params={'PPSX': ppsx,
                                 'PPFT': ppft,
                                 'type': "11",
                                 'NewUser': "1",
                                 'i1': "0",
                                 'i2': "0",
                                 'login': login,
                                 'passwd': passwd}).read()
        # get redirection url
        try:
            self.redirurl = self.findvar(
                r, 'redirurl', 'window\.location\.replace\(\"(.*)\"\)', re.I)
        except ValueError, e:
            if e == "NotFoundVar":
                return False
        return True

    def checklogin(self):
        ''' check if we are logged in '''
        r = self.getpage(self.redirurl)
        # check if we are getting redirected
        try:
            self.redirurl = self.findvar(
                r.read(), 'redirurl', 'window\.location\.replace\(\"(.*)\"\)', re.I, False)
            logger.debug("Found redirection page, fetching")
            r = self.getpage(self.redirurl)
        except ValueError, e:
            pass

        # see what page we got redirected to
        if re.search('login.live.com', r.geturl()):
            logger.error("Redirect back to login page, unable to login: %s",
                         r.geturl())
            return False
        logger.info("Succesful logged in!")
        self.pageurl = r.geturl()
        try:
            self.pageurl = self.findvar(self.pageurl, 'BrowserSupport', 'BrowserSupport.*?targetUrl=(.*)\&', logerror=False)
            self.pageurl = urllib2.unquote(self.pageurl)
            logger.debug("Found BrowserSupport url, fixed url")
        except ValueError, e:
            pass
        logger.debug('main url: %s', self.pageurl)

        # fetch main page and store it for future functions
        r = self.getpage(self.pageurl)
        self.pageurl = r.geturl()
        self.content = r.read()
        # base url
        self.baseurl = self.pageurl[0:self.pageurl.find("/mail/") + 6]

        return True

    def dologin(self):
        ''' Login to Outlook using password'''
        if 'Password' in outlook.configs:
            if outlook.login() and outlook.checklogin():
                return True
            else:
                logger.info("Incorrect username/password, bye bye")
                return False

        logger.info("No password available, bye bye")
        return False

    def getfolders(self):
        ''' reads folder '''
        # read history from disk
        self.status = shelve.open(self.configs['StatusFile'], writeback=True)
        downloaded = []
        if 'Downloaded' in self.configs and not 'Downloaded' in self.status:
            with open(self.configs['Downloaded']) as f:
                # old status file
                for line in f:
                    if re.match('^(.{8}-.{4}-.{4}-.{4}-.{12})$', line):
                        convert = True
                        downloaded.append(line.rstrip().upper())
            if len(downloaded) > 0:
                # convert old downloaded object from GetLive
                logger.info('Converted old Downloaded file %s to Status file' %
                            self.configs['Downloaded'])
        self.status['Downloaded'] = downloaded
        self.status.sync()

        # find SessionId and AuthUser
        self.sessionid = self.findvar(
            self.content, 'SessionId', 'SessionId: "(.*?)"', re.M)
        self.authuser = self.findvar(
            self.content, 'AuthUser', 'AuthUser: "(\d+)"', re.M)

        # get all folders
        folders = self.findvar(self.content, 'folderViewModel',
                               'folderViewModel:\[(.*?)\]', re.M)
        folderlist = re.findall('{.*?}', folders)

        if not 'folders' in self.status:
            self.status['folders'] = {}

        for l in folderlist:
            g = re.search("fid:'(.*?)',name:'(.*?)',count:(\d+)", l)
            if g is None:
                logger.error('Unable to match fid/name/count: %s', l)
                return False
            fid = g.group(1)
            name = g.group(2).decode('unicode-escape')
            count = g.group(3)
            if not fid in self.status['folders']:
                self.status['folders'][fid] = {'name': name, 'count': count,
                                               'available': [], 'downloaded': [], 'foundall': False}
            else:
                # update
                self.status['folders'][fid]['name'] = name
                self.status['folders'][fid]['count'] = count

            logger.info('Folder %s - %s', fid, name)

        return True

    def getmessageids(self):
        ''' retrieve message ids needed for downloading messages '''
        # get message ids from each folder
        # from now on, add the mt header to each ajax request
        self.br.addheaders = [('mt', self.getcookie('mt'))]
        ajaxpageurl = "%smail.fpp?cnmn=Microsoft.Msn.Hotmail.Ui.Fpp.MailBox.GetInboxData&ptid=0&a=%s&au=%s" % (self.baseurl, self.sessionid, self.authuser)

        for fid in self.status['folders']:
            logger.info(
                'Processing Folder %s', self.status['folders'][fid]['name'])
            d = 'true,false,true,{"%s",null,null,FirstPage,5,1,null,null,null,Date,false,false,null,null,0,Off,-1,null,null,false},false,null' % (fid)
            r = self.getpage(ajaxpageurl,
                             params={
                             'cn': 'Microsoft.Msn.Hotmail.Ui.Fpp.MailBox',
                             'd': d,
                             'mn': 'GetInboxData',
                             'v': 1})

            msgbody = r.read()
            stillpagestogo = True
            messages = []
            oldmessages = 0
            pagenr = 1

            while(stillpagestogo):
                lastID = ''
                lastTZ = ''
                # determine number of messages
                m = self.findvar(msgbody, 'msgtoscan/msgbody', r'messageListPane.*?mCt=\\"(\d+)\\"(.*)', re.I)

                msgtoscan = m[0]
                msgbody = m[1]

                # find all messages on this page
                fi = re.finditer(r'<li class=\\"(.*?)\\" id=\\"([0-9a-f-]{36})\\(.*?)"(msg|conv).*?mdt=\\"(.*?)\\".*?<span email=\\"(.*?)\\".*?<a href=.*?>(.*?)<\/a>', msgbody, re.I)
                for msg in fi:
                    readindicator = msg.group(1)
                    msgid = msg.group(2).upper()
                    style = msg.group(3)
                    msgorconf = msg.group(4)
                    lastID = msgid
                    lastTZ = msg.group(5)
                    From = self.htmlparser.unescape(msg.group(6))
                    Subject = msg.group(7)
                    if Subject.endswith('&#x200f;'):  # why is every subject ending with RTL char?
                        Subject = self.htmlparser.unescape(Subject[:-8])
                    if re.search('conv', msgorconf):
                        logger.error("Conversation mode detected, please setup your mail without conversations")
                        return False
                    else:
                        messages.append(msgid)
                        logger.info('Message From %s Subject %s Read %s' %
                                    (From, Subject, readindicator))

                self.status['folders'][fid]['available'].extend(messages)
                # update folder structure and sync to disk!
                self.status.sync()

                if not re.search('"mlPageNav.*?EndOfList', msgbody):
                    # there is a next page, do we want to go?. If we have
                    # everything for this page, lets stop now
                    contnextpage = False  # default we do not continue
                    for msgid in messages:
                        if not msgid in self.status['Downloaded'] and not msgid in self.status['folders'][fid]['downloaded'] and not msgid in self.status['folders'][fid]['available']:
                            logger.debug('Message %s is new' % (msgid))
                            oldmessages = 0
                            # always continue one page if we found a new
                            # message on this page
                            contnextpage = True
                        else:
                            oldmessages += 1

                    if not contnextpage and 'BreakOnAlreadyDownloaded' in self.configs and int(self.configs['BreakOnAlreadyDownloaded']) > 0 and oldmessages > int(self.configs['BreakOnAlreadyDownloaded']):
                        logger.info('Stop scanning for this folder')
                        contnextpage = False
                    else:
                         contnextpage = True

                    if not self.status['folders'][fid]['foundall'] or contnextpage:  # find out if we need to fetch the page in the first place
                        pagenr += 1
                        logger.info("Fetching next page %s", pagenr)
                        d = 'true,false,true,{"%s",null,,2,5,%s,"%s","%s",null,Date,false,false,"",null,0,Off,%s,null,null,false},false,null' % (fid, pagenr, lastID, lastTZ, msgtoscan)
                        r = self.getpage(ajaxpageurl,
                                         params={
                                         'cn': 'Microsoft.Msn.Hotmail.Ui.Fpp.MailBox',
                                         'd': d,
                                         'mn': 'GetInboxData',
                                         'v': 1})
                        msgbody = r.read()
                        stillpagestogo = True
                        messages = []
                    else:
                        stillpagestogo = False
                else:
                    stillpagestogo = False
                    self.status['folders'][fid]['foundall'] = True
        return True

    def downloadmessages(self):
        ''' download all messages '''
        logger.info('Downloading messages')
        for fid in self.status['folders']:
            foldername = self.status['folders'][fid]['name']
            logger.info('Downloading messages from folder %s' % foldername)
            for msga in self.status['folders'][fid]['available'][:]:
                # see if we have this msg in global (old) Downloaded or
                # downloaded
                if msga.upper() in (msg.upper() for msg in self.status['folders'][fid]['downloaded']):
                    logger.debug('Already downloaded msg %s' % msga)
                    self.status['folders'][fid]['available'].remove(msga)
                elif msga.upper() in (msg.upper() for msg in self.status['Downloaded']):
                    logger.debug('Already downloaded msg %s' % msga)
                    self.status['folders'][fid][
                        'downloaded'].append(msga.upper())
                    self.status['folders'][fid]['available'].remove(msga)
                    self.status['Downloaded'].remove(msga.upper())
                else:
                    logger.info('Downloading message %s' % msga)
                    msg = self.downloadmessage(msga, foldername)
                    if msg:
                        with open("%s/%s" % (self.configs['DestinationDir'], foldername), "a+") as f:
                            f.write(msg)

                        self.status['folders'][fid]['downloaded'].append(
                            msga.upper())
                        self.status['folders'][fid]['available'].remove(msga)

                self.status.sync()
        return True

    def downloadmessage(self, msgidx, foldername):
        ''' dowloads one message and returns the converted mbox-style mail '''
        pageurl = "%sGetMessageSource.aspx?msgid=%s" % (self.baseurl, msgidx)
        r = self.getpage(pageurl)
        messageblock = r.read()

        try:
            pre = self.findvar(messageblock, 'messageblock', "<pre>(.*)</pre>")
        except ValueError:
            return None
        try:
            unescapedmsg = self.htmlparser.unescape(pre).encode('latin1')
        except:
            logger.error("Unable to unescape html of message\n%s", pre)
            return None
        # create a message object to convert it to mbox format
        try:
            msg = email.message_from_string(unescapedmsg)
        except:
            logger.error(
                "Unable to create message object from text\n%s", unescapedmsg)
        # add headers
        msg.add_header("X-GetOutlook-Version", self.version())
        msg.add_header("X-GetOutlook-msgidx", msgidx)
        msg.add_header("X-GetOutlook-Folder", foldername)
        # make flat
        msg_out = StringIO()
        msg_gen = Generator(msg_out, mangle_from_=True)
        msg_gen.flatten(msg, unixfrom=True)
        return msg_out.getvalue()


outlook = Outlook()
outlook.parseargs() and outlook.dologin() and outlook.getfolders() and outlook.downloadmessages() and outlook.getmessageids() and outlook.downloadmessages()
