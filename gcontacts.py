#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""Synchronize google contacts"""

import sys, os.path, shutil
import urllib, urllib2
import xml2vcf, vcf2xml
import vobject, codecs, locale
import shelve, logging
from datetime import datetime
from configobj import ConfigObj
from optparse import OptionParser

__scriptname__ = 'Contacts-Sync'
__version__ = '0.1alpha'
_encoding = locale.getpreferredencoding()
_dtformat = '%Y-%m-%dT%H:%M:%S.%fZ'
options = ConfigObj(os.path.expanduser('~/.gsyncrc'))
contactdb = shelve.open(os.path.expanduser('~/.gcontactsdb'), writeback=True)
# TODO: if writeback is detrimental to performance, try to workaround
# parse command line options
usage = 'usage: %prog [options]'
parser = OptionParser(usage=usage, version='%prog ' + __version__)
parser.add_option('-f', '--force-all', dest='getall', action='store_true',
        help='download and compare all contacts, not just those changed since last run',
        default=False)
parser.add_option('-r', '--remote', dest='preferlocal', action='store_false',
        help='prefer remote if contacts differ (overrides config file)',
        default=None)
parser.add_option('-l', '--local', dest='preferlocal', action='store_true',
        help='prefer local if contacts differ (overrides config file)',
        default=None)
parser.add_option('-v', '--verbose', dest='verbose', action='store_true',
        help='print what\'s happening (loglevel = debug)')
(runoptions, args) = parser.parse_args()

# set logging
if 'loglevel' in options.keys():
    if options['loglevel'].upper() not in logging._levelNames:
        options['loglevel'] = 'debug'
else:
    options['loglevel'] = 'debug'

def makelogger(loglevel):
    logger = logging.getLogger(__scriptname__)
    logger.setLevel(loglevel)
    ch = logging.StreamHandler()
    ch.setLevel(loglevel)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s',
            '%H:%M:%S')
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

if runoptions.verbose:
    logger = makelogger(logging.DEBUG)
else:
    logger = makelogger(logging._levelNames[options['loglevel'].upper()])

def authenticate(user, passwd):
    """attempt to authenticate user"""
    def failed(msg):
        logger.critical(u'Authentication failed: ' + msg)
        sys.exit(2)
    data = {'Email': user, 'Passwd': passwd, 'accountType': 'GOOGLE',
            'source': __scriptname__, 'service': 'cp'}
    datastring = urllib.urlencode(data)
    try:
        gdata = urllib2.urlopen('https://www.google.com/accounts/ClientLogin',
                datastring)
    except urllib2.HTTPError, msg:
        failed(msg)
    gdatatext = gdata.readlines()
    gdata.close()
    try:
        auth = gdatatext[2]
    except IndexError:
        failed('unexpected response')
    if auth[0:5] != 'Auth=':
        failed('unexpected response')
    # change first 'A' to 'a'
    return 'a' + auth[1:].strip()

def getcontacts(user, auth, contactid=None, data=None):
    """download contacts from google
        use data dictionary for:
          - updated-min
          - max-results
          - start-index
          - orderby
          - showdeleted
          - requirealldeleted
          - sortorder
          - group
        http://code.google.com/apis/contacts/docs/3.0/reference.html#Parameters
    """
    # construct header
    headers = {'Authorization': 'GoogleLogin ' + auth,
            'GData-Version': '3.0'}
    url = 'http://www.google.com/m8/feeds/contacts/' + user + '/full'
    if contactid:
        # add contact UID (end of id url) to url to specify single contact
        # http://code.google.com/apis/contacts/docs/3.0/developers_guide_protocol.html
        #   #retrieving_single_contact
        url += '/' + contactid
    elif data:
        # can't send data as POST, append to url
        url += '?' + urllib.urlencode(data)
    gcontactsrequest = urllib2.Request(url, None, headers)
    try:
        # Error 404: Not Found if contact has been deleted
        gcontactsconn = urllib2.urlopen(gcontactsrequest)
    except urllib2.HTTPError, msg:
        handleconnectionerror(msg)
    # convert to unicode
    gencoding = False
    for h in gcontactsconn.headers['content-type'].split(';'):
        if 'charset' in h:
            gencoding = h.split('=')[1]
    if not gencoding:
        gencoding = _encoding
    gcontacts = gcontactsconn.read()
    gcontactsconn.close()
    return unicode(gcontacts, gencoding)

def sendcontact(user, auth, contactxml, contactid=None, delete=False):
    """send new or edited contact to google, or delete existing"""
    # construct header
    # TODO: replace X-HTTP-Method-Override header with proper request from httplib
    #       'X-HTTP-Method-Override': 'POST',
    headers = {'Authorization': 'GoogleLogin ' + auth,
            'Content-Type': 'application/atom+xml',
            'GData-Version': '3.0'}
    if delete:
        # TODO: not sure if deletion is working yet.
        if not contactid:
            logger.error(u'Cannot delete unspecified contact.')
            return False
        headers['X-HTTP-Method-Override'] = 'DELETE'
        del headers['Content-Type']
        contactxml = None
    url = 'http://www.google.com/m8/feeds/contacts/' + user + '/full'
    if contactid:
        # use PUT to update existing contact
        url += '/' + contactid
        headers['X-HTTP-Method-Override'] = 'PUT'
    gcontactsrequest = urllib2.Request(url, contactxml, headers)
    # this will fail with Error 412: Precondition Failed if sent contact exists
    # and has different etag - i.e. has been changed on Google since last sync.
    # Error 404: Not Found if contact has been deleted
    # TODO: handle this somehow
    # perhaps make sure we want to overwrite then delete + add new
    try:
        gcontactsconn = urllib2.urlopen(gcontactsrequest)
    except urllib2.HTTPError, msg:
        logger.error(url)
        logger.error(headers)
        logger.error(contactxml)
        handleconnectionerror(msg)
    # convert to unicode
    gencoding = False
    for h in gcontactsconn.headers['content-type'].split(';'):
        if 'charset' in h:
            gencoding = h.split('=')[1]
    if not gencoding:
        gencoding = _encoding
    gcontacts = gcontactsconn.read()
    gcontactsconn.close()
    return unicode(gcontacts, gencoding)

def handleconnectionerror(msg):
    """ deal with errors from google connection """
    # TODO: not sure why these come as errors instead of response codes...
    # see http://code.google.com/apis/gdata/docs/2.0/reference.html#HTTPStatusCodes
    logger.error(msg)

def comparevcards(vcard, localvcard, auth):
    """ look for local version of this vcard and compare
    should return tuple(action, xml, id, name) """
    id = vcard.uid.value
    name = vcard.fn.value
    # compare
    if vcard.serialize() == localvcard.serialize():
        return localvcard
    # compare REV strings
    if 'rev' in vcard.contents and 'rev' in localvcard.contents:
        # TODO google returns utc times - should make this timezone aware
        vcardrev = datetime.strptime(vcard.rev.value, _dtformat)
        localvcardrev = datetime.strptime(localvcard.rev.value, _dtformat)
        logger.debug(u'Comparing revision times: R{0}, L{1}.'.format(
                vcard.rev.value, localvcard.rev.value))
        if vcardrev > localvcardrev:
            # write new local vcard
            logger.info(u'Local version of contact "{0}" updated.'.format(name))
            return vcard
        elif localvcardrev > vcardrev:
            # update remote
            xmlobj = vcf2xml.toXml(localvcard)
            xml = vcf2xml.ET.tostring(xmlobj, encoding=_encoding)
            response = sendcontact(options['user'], auth, xml, id)
            localvcard = xml2vcf.readXml(response)[0]
            logger.info(u'Remote version of contact "{0}" updated.'.format(name))
            return localvcard
        else:
            logger.debug(u'Revision times are equal.')
    else:
        # use default resolution
        logger.debug(u'One or both versions missing revision time, ' +
                'choosing default resolution method.')
    r = u'Contact "{0}" differs: '.format(name)
    if options['defaultresolution'] == 'prefer local' \
            or runoptions.preferlocal is True:
        xmlobj = vcf2xml.toXml(localvcard)
        xml = vcf2xml.ET.tostring(xmlobj, encoding=_encoding)
        response = sendcontact(options['user'], auth, xml, id)
        localvcard = xml2vcf.readXml(response)[0]
        logger.info(r + u'remote version updated.')
        return localvcard
    elif options['defaultresolution'] == 'prefer remote' \
            or runoptions.preferlocal is False:
        logger.info(r + u'local version updated.')
        return vcard
    elif options['defaultresolution'] == 'do nothing':
        logger.warning(r + u'unable to resolve.')
        versions = u'Local version:\n{0}Remoteversion:\n{1}'.format(
                unicode(localvcard.serialize(), _encoding),
                unicode(vcard.serialize(), _encoding))
        logger.warning(versions)
        return localvcard

def getlocalcontacts():
    """ open local contacts, return dict of (uid, contact) """
    # open contacts file
    logger.debug(u'Opening local contacts file.')
    contactsfilename = os.path.expanduser(options['contacts'])
    contactsfile = codecs.open(contactsfilename, 'r', _encoding)
    contactstext = contactsfile.read()
    contactsfile.close()
    # copy file to archive version
    # TODO do this smarter and optionally
    archivefilename = contactsfilename + '.' + \
            datetime.now().strftime(_dtformat)
    shutil.copy(contactsfilename, archivefilename)
    # parse into dictionary of contacts
    logger.debug(u'Parsing contacts.')
    def getOrMakeUid(c):
        if 'uid' in c.contents:
            return (c.uid.value, c)
        else:
            return (c.fn.value, c)
    localcontacts = dict([getOrMakeUid(c) for c
            in vobject.readComponents(contactstext)])
    return localcontacts

def getlocalchanges(localcontacts):
    # make list of recently changed local contact ids
    localchanges = []
    if 'lastsync' in contactdb.keys():
        lastsync = datetime.strptime(contactdb['lastsync'], _dtformat)
        logger.debug(u'Looking for changes since {0}'.format(contactdb['lastsync']))
        for cuid, c in localcontacts.items():
            if 'rev' in c.contents:
                crev = datetime.strptime(c.rev.value, _dtformat)
                if crev > lastsync:
                    localchanges.append(cuid)
    if len(localchanges):
        logger.debug(u'Detected {0} local changes.'.format(len(localchanges)))

    # make list of recently deleted local contact ids
    if 'cuids' in contactdb.keys():
        logger.debug(u'Looking for additions and deletions since last sync.')
        localadditions = set(localcontacts.keys()) - set(contactdb['cuids'])
        localdeletions = set(contactdb['cuids']) - set(localcontacts.keys())
    else:
        localadditions = localcontacts.keys()
        localdeletions = []
    if len(localadditions):
        logger.debug(u'Detected {0} local additions.'.format(len(localadditions)))
    if len(localdeletions):
        logger.debug(u'Detected {0} local deletions.'.format(len(localdeletions)))

    return list(localadditions), localchanges, list(localdeletions)

def getallfromgoogle():
    """ get all contacts from google and save to file """
    auth = authenticate(options['user'], options['password'])
    data = {'max-results': '2000'}
    contactsxml = getcontacts(options['user'], auth, data=data)
    return contactsxml

def savexml(xml):
    xmlfilename = 'google-contacts-{0}.xml'.format(datetime.now().strftime(_dtformat))
    xmlfile = codecs.open(xmlfilename, 'w', _encoding)
    xmlfile.write(xml)
    xmlfile.close()

def execute():
    localcontacts = getlocalcontacts()
    localadditions, localchanges, localdeletions = getlocalchanges(localcontacts)

    # get (recently changed) contacts from google
    data = {'max-results': len(localcontacts) + 50}
    if 'lastsync' in contactdb.keys() and not runoptions.getall:
        data['updated-min'] = contactdb['lastsync']
    logger.debug(u'Logging into Google Contacts.')
    auth = authenticate(options['user'], options['password'])
    logger.debug(u'Retrieving contact list.')
    contactsxml = getcontacts(options['user'], auth, data=data)
    # store xml for reference
    if options['loglevel'] == 'debug':
        savexml(contactsxml)

    # parse into individual vcards
    logger.debug(u'Parsing contacts from Google.')
    contacts = xml2vcf.readXml(contactsxml)
    logger.info(u'Received {0} contacts from Google.'.format(len(contacts)))
    for c in contacts:
        if c.uid.value in localdeletions:
            # don't bother comparing if we're going to delete it anyway
            logger.debug(u'Ignoring "{0}": in deletion list.'.format(c.fn.value))
            continue

        if c.uid.value in localcontacts:
            #logger.debug(u'Comparing "{0}".'.format(c.fn.value))
            localcontacts[c.uid.value] = comparevcards(c,
                    localcontacts[c.uid.value], auth)
        elif 'fn' in c.contents:
            localcontacts[c.uid.value] = c
            logger.info(u'New contact "{0}" added.'.format(c.fn.value))
        else:
            logger.debug(u'New unparseable remote contact ' +
                    u'"{0}" ignored.'.format(c.uid.value))

        if c.uid.value in localchanges:
            # already compared, so delete from localchanges list
            del localchanges[localchanges.index(c.uid.value)]

    # local additions
    # TODO: additions go to general contact list, not My Contacts, and have to
    # be moved manually in Gmail. Fix this.
    logger.debug(u'Sending local additions.')
    for n in localadditions:
        xmlobj = vcf2xml.toXml(localcontacts[n])
        xml = vcf2xml.ET.tostring(xmlobj, encoding=_encoding)
        response = sendcontact(options['user'], auth, xml)
        localvcard = xml2vcf.readXml(response)[0]
        # replace original with uid/etagged version from google
        del localcontacts[n]
        localcontacts[localvcard.uid.value] = localvcard
        logger.info(u'Local contact "{0}" added to Google.'.format(n))

    # TODO: deal with remote deletions
    # remote deletions should have only <atom:id> and <gd:deleted> for 30 days
    # need special query? xml2vcf.readXml might fail as it's invalid vcard
    # without N, FN
    for cuid in localdeletions:
        #logger.debug(u'Deleting "{0}": deleted locally.'.format(c.fn.value))
        #response = sendcontact(options['user'], auth, contactxml, contactid, True)
        # TODO: deal with response
        # sendcontact returns False if no contactid specified
        # add to list to delete until this is worked out
        logger.debug(u'Recording deletion of contact {0}.'.format(cuid))
        try:
            contactdb['todelete'].append(cuid)
            # NB Shelf.append() is dependent on writeback=True
        except KeyError:
            contactdb['todelete'] = [cuid]
        pass

    # remaining localchanges
    logger.debug(u'Examining local changes.')
    for cuid in localchanges:
        logger.debug(u'Retrieving "{0}".'.format(localcontacts[cuid].fn.value))
        contactxml = getcontacts(options['user'], auth, cuid)
        logger.debug(u'Parsing contact from Google.')
        contact = xml2vcf.readXml(contactxml)[0]
        logger.debug(u'Comparing "{0}".'.format(localcontacts[cuid].fn.value))
        localcontacts[cuid] = comparevcards(contact, localcontacts[cuid], auth)

    # write out contacts file
    logger.debug(u'Writing new local file.')
    contactsfile = codecs.open(os.path.expanduser(options['contacts']), 'w',
            _encoding)
    # sort list by cuid so diffs are easier
    for cuid, c in sorted(localcontacts.items()):
        contactsfile.write(c.serialize().decode(_encoding))
        # should this be?
        #contactsfile.write(unicode(c.serialize(), _encoding))
    contactsfile.close()

    # set last sync time in config: now() or utcnow()?
    logger.debug(u'Recording sync details.')
    contactdb['lastsync'] = datetime.strftime(datetime.utcnow(), _dtformat)
    contactdb['cuids'] = localcontacts.keys()
    contactdb.close()

if __name__ == '__main__':
    execute()

