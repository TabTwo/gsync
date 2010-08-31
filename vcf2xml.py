#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""Convert vcard to xml for google contacts api"""

import xml.etree.cElementTree as ET
import vobject, codecs, locale
import sys, os.path
from optparse import OptionParser

__version__ = '0.1alpha'
_encoding = locale.getpreferredencoding()

# strings for google
namespaces = {
        'atom':       'http://www.w3.org/2005/Atom',
        'gd':         'http://schemas.google.com/g/2005',
        'gcontact':   'http://schemas.google.com/contact/2008',
        'batch':      'http://schemas.google.com/gdata/batch',
        'opensearch': 'http://a9.com/-/spec/opensearch/1.1/' }
# TODO look at ET._namespace_map dictionary
idbase = 'http://www.google.com/m8/feeds/contacts/mark.opus11%40googlemail.com/base/'
grbase = 'http://www.google.com/m8/feeds/groups/mark.opus11%40googlemail.com/base/'
# TODO make these strings from user name
orgRelTypes = {
        'WORK': 'http://schemas.google.com/g/2005#work' }
#       'OTHER': 'http://schemas.google.com/g/2005#other' }
emailRelTypes = {
        'HOME': 'http://schemas.google.com/g/2005#home',
        'WORK': 'http://schemas.google.com/g/2005#work' }
#       'OTHER': 'http://schemas.google.com/g/2005#other' }
addressRelTypes = emailRelTypes
phoneRelTypesPairs = {
        ('COMPANY', 'MAIN'): 'http://schemas.google.com/g/2005#company_main',
        ('HOME', 'FAX'): 'http://schemas.google.com/g/2005#home_fax',
        ('OTHER', 'FAX'): 'http://schemas.google.com/g/2005#other_fax',
        ('TTY', 'TTD'): 'http://schemas.google.com/g/2005#tty_tdd',
        ('WORK', 'FAX'): 'http://schemas.google.com/g/2005#work_fax',
        ('WORK', 'CELL'): 'http://schemas.google.com/g/2005#work_mobile',
        ('WORK', 'PAGER'): 'http://schemas.google.com/g/2005#work_pager' }
phoneRelTypes = {
        'ASSISTANT': 'http://schemas.google.com/g/2005#assistant',
        'CALLBACK': 'http://schemas.google.com/g/2005#callback',
        'CAR': 'http://schemas.google.com/g/2005#car',
        'FAX': 'http://schemas.google.com/g/2005#fax',
        'HOME': 'http://schemas.google.com/g/2005#home',
        'ISDN': 'http://schemas.google.com/g/2005#isdn',
        'MAIN': 'http://schemas.google.com/g/2005#main',
        'CELL': 'http://schemas.google.com/g/2005#mobile',
        'PAGER': 'http://schemas.google.com/g/2005#pager',
        'RADIO': 'http://schemas.google.com/g/2005#radio',
        'TELEX': 'http://schemas.google.com/g/2005#telex',
        'WORK': 'http://schemas.google.com/g/2005#work' }
#       'OTHER': 'http://schemas.google.com/g/2005#other',

def execute():
    usage = 'usage: %prog [input file] | xmllint --format - > out.xml'
    parser = OptionParser(usage=usage, version='%prog ' + __version__)
    (options, args) = parser.parse_args()

    # open vcard file
    try:
        vcardfile = codecs.open(args[0], 'r', _encoding)
    except IOError, msg:
        print >>sys.stderr, msg
        return 2
    vcard = vcardfile.read()
    vcardfile.close()
    # TODO: this is almost useless since Google won't accept a feed of entries.
    # Should perhaps adapt this to output a batch document.
    # http://code.google.com/apis/gdata/docs/batch.html
    feed = ET.Element(addNS('feed', 'atom'))
    for i, vc in enumerate(vobject.readComponents(vcard)):
        xml = toXml(vc)
        feed.insert(i, xml)
    print ET.tostring(feed, encoding=_encoding)

def addNS(tag, namespace):
    """ add a namespace from the namespace dictionary to a tag """
    return '{{{0}}}{1}'.format(namespaces[namespace], tag)

def stripAndJoin(string, removeNewLines=True):
    """ remove \n, \r, double spaces from string """
    if removeNewLines:
        string = string.replace('\n', '').replace('\r', '')
    while '  ' in string:
        string = string.replace('  ', ' ').strip()
    return string

def toXml(vcard):
    """ convert a vcard to xml format """
    xml = ET.Element(addNS('entry', 'atom'))
    # add google category element
    cat = ET.SubElement(xml, addNS('category', 'atom'),
            scheme='http://schemas.google.com/g/2005#kind',
            term='http://schemas.google.com/contact/2008#contact')

    # convert each item from vcard
    for component in vcard.getChildren():
        if component.name == u'X-GOOGLE-ETAG':
            addEtag(xml, component)
        if component.name == u'UID':
            addId(xml, component)
        elif component.name == u'REV':
            addUpdated(xml, component)
        elif component.name == u'N':
            addName(xml, component)
        elif component.name == u'ORG':
            addOrganization(xml, component)
        elif component.name == u'ROLE':
            addJobDescription(xml, component)
        elif component.name == u'TITLE':
            addOrgTitle(xml, component)
        elif component.name == u'NOTE':
            addContent(xml, component)
        elif component.name == u'EMAIL':
            addEmail(xml, component)
        elif component.name == u'TEL':
            addPhoneNumber(xml, component)
        elif component.name == u'ADR':
            addAddress(xml, component)
        elif component.name == u'X-GOOGLE-GROUP':
            addGroup(xml, component)

    return xml

def addEtag(xml, component):
    etag = '"{0}."'.format(component.value)
    xml.set(addNS('etag', 'gd'), etag)

def addId(xml, component):
    # actual id has url at start
    id = ET.SubElement(xml, addNS('id', 'atom'))
    id.text = idbase + component.value

def addGroup(xml, component):
    group = ET.SubElement(xml, addNS('groupMembershipInfo', 'gcontact'))
    group.set('deleted', 'false')
    group.set('href', grbase + component.value)

def addUpdated(xml, component):
    updated = ET.SubElement(xml, addNS('updated', 'atom'))
    updated.text = component.value

def addName(xml, component):
    name = ET.SubElement(xml, addNS('name', 'gd'))
    fullname = ['', '', '', '', '']
    for t, v in component.value.__dict__.items():
        if v == '':
            continue
        if t == u'prefix':
            n = ET.SubElement(name, addNS('namePrefix', 'gd'))
            n.text = v
            fullname[0] = v
        elif t == u'given':
            n = ET.SubElement(name, addNS('givenName', 'gd'))
            n.text = v
            fullname[1] = v
        elif t == u'additional':
            n = ET.SubElement(name, addNS('additionalName', 'gd'))
            n.text = v
            fullname[2] = v
        elif t == u'family':
            n = ET.SubElement(name, addNS('familyName', 'gd'))
            n.text = v
            fullname[3] = v
        elif t == u'suffix':
            n = ET.SubElement(name, addNS('nameSuffix', 'gd'))
            n.text = v
            fullname[4] = v
    # add full name
    n = ET.SubElement(name, addNS('fullName', 'gd'))
    n.text = stripAndJoin(' '.join(fullname))

def addOrganization(xml, component):
    # check if there is already an organization element
    # (may have been added by ROLE or TITLE)
    organization = xml.find(addNS('organization', 'gd'))
    if not organization:
        organization = ET.SubElement(xml, addNS('organization', 'gd'))
    # deal with types
    relorlabel = None
    try:
        for t in component.type_paramlist:
            if t == u'PREF':
                organization.set('primary', 'true')
            elif t in orgRelTypes and relorlabel is None:
                organization.set('rel', orgRelTypes[t])
                relorlabel = 'rel'
            elif relorlabel is None:
                organization.set('label', t.capitalize())
                relorlabel = 'label'
    except AttributeError:
        pass
    if relorlabel is None:
        # organization must have exactly one rel or label
        organization.set('rel', orgRelTypes['WORK'])
    # value
    orgname = ET.SubElement(organization, addNS('orgName', 'gd'))
    orgname.text = component.value[0]
    if len(component.value) > 1:
        orgdepartment = ET.SubElement(organization, addNS('orgDepartment', 'gd'))
        orgdepartment.text = component.value[1]

def addJobDescription(xml, component):
    # check if there is already an organization element
    # (may have been added by ORG or TITLE)
    organization = xml.find(addNS('organization', 'gd'))
    if not organization:
        organization = ET.SubElement(xml, addNS('organization', 'gd'))
    jd = ET.SubElement(organization, addNS('orgJobDescription', 'gd'))
    jd.text = component.value
    if organization.get('rel') is None and organization.get('label') is None:
        # organization must have exactly one rel or label
        organization.set('rel', orgRelTypes['WORK'])

def addOrgTitle(xml, component):
    # check if there is already an organization element
    # (may have been added by ORG or ROLE)
    organization = xml.find(addNS('organization', 'gd'))
    if not organization:
        organization = ET.SubElement(xml, addNS('organization', 'gd'))
    tit = ET.SubElement(organization, addNS('orgTitle', 'gd'))
    tit.text = component.value
    if organization.get('rel') is None and organization.get('label') is None:
        # organization must have exactly one rel or label
        organization.set('rel', orgRelTypes['WORK'])

def addContent(xml, component):
    note = ET.SubElement(xml, addNS('content', 'atom'), type='text')
    note.text = stripAndJoin(component.value)

def addEmail(xml, component):
    email = ET.SubElement(xml, addNS('email', 'gd'))
    email.set('address', component.value)
    # deal with types
    relorlabel = None
    try:
        for t in component.type_paramlist:
            if t == u'PREF':
                email.set('primary', 'true')
            elif t == u'INTERNET':
                continue
            elif t in emailRelTypes and relorlabel is None:
                email.set('rel', emailRelTypes[t])
                relorlabel = 'rel'
            elif relorlabel is None:
                email.set('label', t.capitalize())
                relorlabel = 'label'
    except AttributeError:
        pass
    if relorlabel is None:
        # email must have exactly one rel or label
        email.set('rel', emailRelTypes['HOME'])

def addPhoneNumber(xml, component):
    phone = ET.SubElement(xml, addNS('phoneNumber', 'gd'))
    phone.text = component.value
    # deal with types
    try:
        types = component.type_paramlist
    except AttributeError:
        types = None
    relorlabel = None
    if types:
        # pairs first
        for a, b in phoneRelTypesPairs.keys():
            if a in types and b in types and relorlabel is None:
                phone.set('rel', phoneRelTypesPairs[(a, b)])
                del types[types.index(a)]
                del types[types.index(b)]
                relorlabel = 'rel'
        # remaining types
        for t in types:
            if t == u'PREF':
                phone.set('primary', 'true')
            elif t in phoneRelTypes and relorlabel is None:
                phone.set('rel', phoneRelTypes[t])
                relorlabel = 'rel'
            elif relorlabel is None:
                phone.set('label', t.capitalize())
                relorlabel = 'label'
    else:
        # phonenumber must have exactly one rel or label
        phone.set('rel', phoneRelTypes['HOME'])

def addAddress(xml, component):
    # check if empty
    if component.value.__str__().strip('\n ,') == '':
        return
    address = ET.SubElement(xml, addNS('structuredPostalAddress', 'gd'))
    # deal with types
    relorlabel = None
    try:
        for t in component.type_paramlist:
            if t == u'PREF':
                address.set('primary', 'true')
            elif t in addressRelTypes and relorlabel is None:
                address.set('rel', addressRelTypes[t])
                relorlabel = 'rel'
            elif relorlabel is None:
                address.set('label', t.capitalize())
                relorlabel = 'label'
    except AttributeError:
        pass
    if relorlabel is None:
        # address mus have exactly one rel or label
        address.set('rel', addressRelTypes['HOME'])
    # values
    fulladdress = [''] * 7
    for t, v in component.value.__dict__.items():
        if v == '':
            continue
        if t == u'box':
            a = ET.SubElement(address, addNS('pobox', 'gd'))
            a.text = v
            fulladdress[0] = v
        elif t == u'extended':
            a = ET.SubElement(address, addNS('housename', 'gd'))
            a.text = v
            fulladdress[1] = v
        elif t == u'street':
            a = ET.SubElement(address, addNS('street', 'gd'))
            a.text = v
            fulladdress[2] = v
        elif t == u'city':
            a = ET.SubElement(address, addNS('city', 'gd'))
            a.text = v
            fulladdress[3] = v
        elif t == u'region':
            a = ET.SubElement(address, addNS('region', 'gd'))
            a.text = v
            fulladdress[4] = v
        elif t == u'code':
            a = ET.SubElement(address, addNS('postcode', 'gd'))
            a.text = v
            fulladdress[5] = v
        elif t == u'country':
            a = ET.SubElement(address, addNS('country', 'gd'))
            a.text = v
            fulladdress[6] = v
    # add formatted address
    a = ET.SubElement(address, addNS('formattedAddress', 'gd'))
    a.text = stripAndJoin(' '.join(fulladdress))

if __name__ == '__main__':
    sys.exit(execute())
