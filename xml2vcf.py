#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""Convert xml from google contacts api to vcard"""

import xml.etree.cElementTree as ET
import vobject, codecs, locale
import sys, os.path
from optparse import OptionParser

__version__ = '0.1alpha'
_encoding = locale.getpreferredencoding()

def execute():
    usage = 'usage: %prog [options] [input file]'
    parser = OptionParser(usage=usage, version='%prog ' + __version__)
    parser.add_option('-d', '--dir', dest='directory', help='output to DIR',
            metavar='DIR', default=None)
    (options, args) = parser.parse_args()

    if options.directory and not os.path.isdir(options.directory):
        # make dir
        try:
            os.mkdir(options.directory)
        except:
            return 2

    # open xml file
    vcards = readXml(args[0], file=True)

    # output vcards
    for v in vcards:
        if options.directory:
            # make filename from uid or full name
            if 'uid' in v.contents:
                filename = '{0}/{1}.vcf'.format(options.directory, v.uid.value)
            else:
                filename = '{0}/{1}.vcf'.format(options.directory,
                        v.fn.value.replace(' ', '_'))
                # TODO: encoding error here ^?
            file = codecs.open(filename, 'w', _encoding)
            file.write(v.serialize().decode(_encoding))
            file.close()
        else:
            print v.serialize()

def readXml(xml, file=False):
    """ open the xml and find the contact entries """
    if file:
        try:
            xmlfile = ET.parse(xml)
        except IOError, msg:
            print >>sys.stderr, msg
            return 2
        xmlcard = xmlfile.getroot()
    else:
        xmlcard = ET.fromstring(xml.encode(_encoding))
    vcards = []

    if splitNS(xmlcard.tag)[1] == u'feed':
        # this is a list of contacts
        # parse metadata at start of feed?
        """ is this useful?
        xml namespaces in <feed>
        <id>user@gmail.com
        <updated>
        <title>
        numerous <link>s
        <author>
        <generator>
        <openSearch>
        """
        for element in xmlcard.getchildren():
            if splitNS(element.tag)[1] == u'entry':
                vcards.append(parseEntry(element))
    elif splitNS(xmlcard.tag)[1] == u'entry':
        # this is a single contact
        vcards.append(parseEntry(xmlcard))
    else:
        # unknown
        print >>sys.stderr, 'No entries found'
        return None

    return vcards

def splitNS(tag):
    """ split namespace from element names """
    if '}' in tag:
        ns, div, nn = tag.partition('}')
    else:
        ns, nn = (None, tag)
    return ns, nn

def gdRel(relstring):
    """ split a gdata rel string and return vcard type """
    # TODO: CELL special case should be somewhere else...
    return relstring.partition('#')[2].replace('mobile', 'cell').upper()

def unEntity(text):
    """ remove incorrect xml entities """
    while unichr(195) in text:
        # get the character after
        i = text.index(unichr(195))
        newtext = text[:i] + unichr(ord(text[i+1]) + 64)
        if len(text) >= i + 2:
            newtext += text[i+2:]
        text = newtext
    return text

def stripAndJoin(string, removeNewLines=True):
    """ remove \n, \r, double spaces from string """
    if removeNewLines:
        string = string.replace('\n', '').replace('\r', '')
    while '  ' in string:
        string = string.replace('  ', ' ').replace('\n ', '\n').strip()
    return unEntity(string)

def parseEntry(entry):
    """ parse an xml contact entry and return it as a vcard """
    vcard = vobject.vCard()

    # get etag if present
    for k, v in entry.attrib.items():
        if splitNS(k)[1] == 'etag':
            addEtag(vcard, v)

    for element in entry.getchildren():
        nn = splitNS(element.tag)[1]
        """ knowingly ignoring these elements:
                app:edited      seems to be the same as <updated>
                category
                title
                link            can be used for photo
                groupMembershipInfo
            can also use X-??? types to retain info if necessary
        """
        if nn == u'id':
            addUid(vcard, element)
        elif nn == u'updated':
            addRev(vcard, element)
        elif nn == u'name':
            addName(vcard, element)
        elif nn == u'organization':
            addOrg(vcard, element)
        elif nn == u'content':
            addNote(vcard, element)
        elif nn == u'email':
            addEmail(vcard, element)
        elif nn == u'phoneNumber':
            addTel(vcard, element)
        elif nn == u'im':
            addIm(vcard, element)
        elif nn == u'structuredPostalAddress':
            addStructuredAddress(vcard, element)
        elif nn == u'postalAddress':
            addUnstructuredAddress(vcard, element)
        elif nn == u'groupMembershipInfo':
            addGroup(vcard, element)

    return vcard

def addEtag(vcard, etag):
    """ add xml etag to a vcard """
    UID = vcard.add('x-google-etag')
    UID.value = etag.strip('".')

def addUid(vcard, element):
    """ add xml id to a vcard """
    UID = vcard.add('uid')
    idurl = stripAndJoin(element.text)
    if '/' in idurl:
        UID.value = idurl.rpartition('/')[2]
    else:
        UID.value = idurl

def addRev(vcard, element):
    """ add xml updated time to a vcard """
    REV = vcard.add('rev')
    REV.value = stripAndJoin(element.text)

def addName(vcard, element):
    """ add an xml name to a vcard """
    N = vobject.vcard.Name()
    FN = None
    for n in element.getchildren():
        nn = splitNS(n.tag)[1]
        if nn == 'familyName':
            N.family = stripAndJoin(n.text)
        elif nn == 'givenName':
            N.given = stripAndJoin(n.text)
        elif nn == 'additionalName':
            N.additional = stripAndJoin(n.text)
        elif nn == 'namePrefix':
            N.prefix = stripAndJoin(n.text)
        elif nn == 'nameSuffix':
            N.suffix = stripAndJoin(n.text)
        elif nn == 'fullName':
            FN = stripAndJoin(n.text)
        else:
            # catch errors
            pass
    # do we have full name?
    if FN is None:
        FN = '{prefix} {given} {additional} {family} {suffix}'.format(**N.__dict__)
        FN = FNstripAndJoin()
    # add to vcard
    vcard.add('n')
    vcard.n.value = N
    vcard.add('fn')
    vcard.fn.value = FN

def addOrg(vcard, element):
    """ add an xml organization to a vcard """
    ORG = vcard.add('org')
    type = []
    for k, v in element.items():
        if k == 'label':
            type.append(v.upper())
        elif k == 'primary' and v == 'true':
            type.append('PREF')
        elif k == 'rel':
            type.append(gdRel(v))
    if len(type) > 0:
        ORG.type_param = type.pop()
        for t in type:
            ORG.type_paramlist.append(t)
    org = ['', '']
    for n in element.getchildren():
        nn = splitNS(n.tag)[1]
        if nn == 'orgDepartment':
            org[1] = stripAndJoin(n.text)
        elif nn == 'orgJobDescription':
            ROLE = vcard.add('role')
            ROLE.value = stripAndJoin(n.text)
        elif nn == 'orgName':
            org[0] = stripAndJoin(n.text)
        elif nn == 'orgTitle':
            TITLE = vcard.add('title')
            TITLE.value = stripAndJoin(n.text)
    # deal with org
    if org[0] == '' and org[1] != '':
        ORG.value = org[1:]
    elif org[0] != '' and org[1] == '':
        ORG.value = org[:1]
    elif org[0] != '' and org[1] != '':
        ORG.value = org

def addNote(vcard, element):
    """ add xml content to a vcard """
    if element.text is not None:
        vcard.add('note')
        #text = stripAndJoin(element.text, removeNewLines=False)
        text = stripAndJoin(element.text)
        vcard.note.value = text.replace('\\', '')

def addEmail(vcard, element):
    """ add an xml email to a vcard """
    EMAIL = vcard.add('email')
    EMAIL.type_param = 'INTERNET'
    for k, v in element.items():
        if k == 'address':
            EMAIL.value = v.lower()
        elif k == 'displayName':
            pass
        elif k == 'label':
            EMAIL.type_paramlist.append(v.upper())
        elif k == 'rel':
            EMAIL.type_paramlist.append(gdRel(v))
        elif k == 'primary' and v == 'true':
            EMAIL.type_paramlist.append('PREF')

def addTel(vcard, element):
    """ add an xml phone number to a vcard """
    TEL = vcard.add('tel')
    type = []
    for k, v in element.items():
        if k == 'label':
            type.append(v.upper())
        elif k == 'rel':
            t = gdRel(v)
            if '_' in t:
                type += t.split('_')
            else:
                type.append(t)
        elif k == 'uri':
            pass
        elif k == 'primary' and v == 'true':
            type.append('PREF')
    if len(type) > 0:
        TEL.type_param = type.pop()
        for t in type:
            TEL.type_paramlist.append(t)
    # add number
    TEL.value = stripAndJoin(element.text)

def addIm(vcard, node):
    """ add an xml im to a vcard """
    # map to X-YAHOO etc
    pass

def addStructuredAddress(vcard, element):
    """ add an xml structured address to a vcard """
    ADR = vobject.vcard.Address()
    for n in element.getchildren():
        nn = splitNS(n.tag)[1]
        if nn == 'pobox':
            ADR.box = stripAndJoin(n.text)
        elif nn == 'housename':
            ADR.extended = stripAndJoin(n.text)
        elif nn == 'street':
            ADR.street = stripAndJoin(n.text)
        elif nn == 'city':
            ADR.city = stripAndJoin(n.text)
        elif nn == 'region':
            ADR.region = stripAndJoin(n.text)
        elif nn == 'postcode':
            ADR.code = stripAndJoin(n.text)
        elif nn == 'country':
            ADR.country = stripAndJoin(n.text)
        elif nn == 'formattedAddress':
            # use this to check for extra fields?
            pass
        else:
            # catch errors
            pass
    A = vcard.add('adr')
    A.value = ADR
    type = []
    for k, v in element.items():
        if k == 'rel':
            type.append(gdRel(v))
        elif k == 'mailClass':
            pass
        elif k == 'usage':
            pass
        elif k == 'label':
            type.append(v.upper())
        elif k == 'primary' and v == 'true':
            type.append('PREF')
    if len(type) > 0:
        A.type_param = type.pop()
        for t in type:
            A.type_paramlist.append(t)

def addUnstructuredAddress(vcard, element):
    """ add an xml unstructured address to a vcard """
    pass

def addGroup(vcard, element):
    """ add google group info to a vcard """
    GROUP = vcard.add('x-google-group')
    for k, v in element.items():
        if k == 'href':
            if '/' in v:
                GROUP.value = v.rpartition('/')[2]
            else:
                GROUP.value = v

if __name__ == '__main__':
    sys.exit(execute())

""" gData => vCard mapping

gd:additionalName                   N[2] Additional Names
gd:city                             ADR[3] City
gd:comments
gd:country                          ADR[6] Country
gd:deleted
gd:email                            EMAIL
gd:entryLink
gd:extendedProperty
gd:familyName                       N[0] Family Name
gd:feedLink
gd:fullName                         FN
gd:geoPt
gd:givenName                        N[1] Given Name
gd:housename                        ADR[1] Extended address
gd:im
gd:money
gd:name                             container for N
gd:namePrefix                       N[3] Honorific Prefix
gd:nameSuffix                       N[4] Honorific Suffix
gd:organization                     ORG
gd:orgDepartment                    ORG[1]
gd:orgJobDescription                ROLE
gd:orgName                          ORG[0]
gd:orgSymbol
gd:orgTitle                         TITLE
gd:originalEvent
gd:phoneNumber                      TEL
gd:pobox                            ADR[0] PO Box
gd:postalAddress                    container for unstructured ADR
gd:postcode                         ADR[5] Postcode
gd:rating
gd:recurrence
gd:recurrenceException
gd:region                           ADR[4] Region
gd:reminder
gd:resourceId
gd:street                           ADR[2] Street
gd:structuredPostalAddress          container for ADR
gd:when
gd:where
gd:who

gContact:billingInformation
gContact:birthday                   BDAY
gContact:calendarLink
gContact:directoryServer
gContact:event
gContact:externalId
gContact:gender
gContact:groupMembershipInfo        X-GOOGLE-GROUP
gContact:hobby
gContact:initials
gContact:jot
gContact:language
gContact:maidenName
gContact:mileage
gContact:nickname                   NICKNAME
gContact:occupation
gContact:priority
gContact:relation
gContact:sensitivity
gContact:shortName
gContact:subject
gContact:systemGroup
gContact:userDefinedField
gContact:website                    URL

updated                             REV
app:edited                          REV
gd:etag                             X-GOOGLE-ETAG
id                                  UID

link rel='.../rel#photo'            PHOTO
content                             NOTE

Attribute mapping

@rel        xs:string               TYPE
@label      xs:string               TYPE
@primary    xs:boolean              TYPE=pref

Not supported in gData:
    N[3] Honorific Prefix
    N[4] Honorific Suffix

Not supported in vCard:
    gd:im
"""
