#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""Synchronize google calendar"""

import os.path
import codecs, locale
import shelve, logging
import subprocess, hashlib
from datetime import date, datetime, timedelta
from dateutil.tz import tzlocal, tzfile
from dateutil import parser as dtparser
from configobj import ConfigObj
from optparse import OptionParser
# python api works well enough for calendar so use it
import gdata.service, gdata.calendar, gdata.calendar.service
import atom, atom.service
# will have to write my own interface to remind

__scriptname__ = 'Calendar-Sync'
__version__ = '0.1alpha'
_encoding = locale.getpreferredencoding()
_qdtformat = '%Y-%m-%dT%H:%M:%S.000Z'
_dtformat = '%Y-%m-%dT%H:%M:%S%z'
_ddformat = '%Y-%m-%d'
# note convolutions to get colon in timezone offset
options = ConfigObj(os.path.expanduser('~/.gsyncrc'))
caldb = shelve.open(os.path.expanduser('~/.gcaldb'), writeback=True)
# parse command line options
usage = 'usage: %prog [options]'
parser = OptionParser(usage=usage, version='%prog ' + __version__)
parser.add_option('-f', '--force-all', dest='getall', action='store_true',
        help='download and compare all events, not just those changed ' +
        'since last run', default=False)
parser.add_option('-r', '--remote', dest='preferlocal', action='store_false',
        help='prefer remote if events differ (overrides config file)',
        default=None)
parser.add_option('-l', '--local', dest='preferlocal', action='store_true',
        help='prefer local if events differ (overrides config file)',
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

class Event():
    """ hold event information """
    def __init__(self, remline):
        fileinfo, remline = remline.strip().split('\n')
        self.remline = remline
        self.linenumber, self.filename = fileinfo.split(' ')
        self.uid = hashlib.md5(remline.encode(_encoding)).hexdigest()
        fields = remline.split(None, 5)
        # set defaults
        tzfilename = '/usr/share/zoneinfo/' + options['timezone']
        if os.path.isfile(tzfilename):
            self.timezone = tzfile(tzfilename)
        else:
            logging.debug('No timezone file {0}. ' + \
                    'Setting to local zone.'.format(options['timezone']))
            self.timezone = tzlocal()
        self.transp = 'OPAQUE'
        self.categories = []
        self.add_date(fields[0])
        self.add_tags(fields[2])
        self.add_times(fields[4], fields[3])
        self.add_body(fields[5])

    def add_date(self, dt):
        """ split date into components """
        (self.year, self.month, self.day) = [int(i) for i in dt.split('/')]

    def add_tags(self, tags):
        """ add tags to event """
        if tags == '*':
            tags = []
        else:
            tags = tags.split(',')
        # parse tags
        for t in tags:
            if '=' in t:
                (k, v) = t.split('=')
                if k == 'TZ':
                    tzfilename = '/usr/share/zoneinfo/' + v
                    if os.path.isfile(tzfilename):
                        self.timezone = tzfile(tzfilename)
                    else:
                        logging.error('No timezone file {0}'.format(v))
                elif k == 'TRANSP':
                    self.transp = v
            else:
                self.categories.append(t)

    def add_times(self, start, duration):
        """ add times to event """
        if start == '*':
            (self.hour, self.minute) = (None, None)
        else:
            (self.hour, self.minute) = divmod(int(start), 60)
        if duration == '*':
            self.duration = None
        else:
            self.duration = int(duration)
        if self.hour:
            self.dtstart = datetime(self.year, self.month, self.day, self.hour,
                    self.minute, tzinfo=self.timezone)
        else:
            self.dtstart = date(self.year, self.month, self.day)
        if self.duration:
            self.dtend = self.dtstart + timedelta(minutes=self.duration)
        else:
            self.dtend = self.dtstart

    def add_body(self, text):
        """ add summary, location, description """
        # split text into lines
        text = text.split(options['remnewlinechar'])
        sumloc = text[0].split(options['remlocation'])
        self.summary = sumloc[0]
        if len(sumloc) == 1:
            self.location = None
        else:
            self.location = sumloc[1]
        # join remaining lines into description
        if len(text) > 1:
            self.description = '\n'.join(text[1:])
        else:
            self.description = None

    def gdatawhen(self):
        """ return gdata When object for this event """
        # google insists on a colon in the timezone offset which we have to add
        # manually
        start = None
        if type(self.dtstart).__name__ == 'date':
            start = self.dtstart.strftime(_ddformat)
        else:
            start = self.dtstart.strftime(_dtformat)
            start = start[:-2] + ':' + start[-2:]
        end = None
        if type(self.dtend).__name__ == 'date':
            end = self.dtend.strftime(_ddformat)
        else:
            end = self.dtend.strftime(_dtformat)
            end = end[:-2] + ':' + end[-2:]
        return gdata.calendar.When(start_time=start, end_time=end)

class Remevent():
    """ a remind event """
    def __init__(self, gevent):
        """ convert a google event to a remind event """
        remstring = 'REM {date}{time}{dur}{tag} \\\n\t' + \
                'MSG %g %3 %"{summary}{location}{description}%"%\n'
        remdateformat = '%b %-d %Y'
        remtimeformat = '%H:%M'
        # format for rem -s
        remsstring = '{date} * {tag} {dur} {time} {body}'
        remsdateformat = '%Y/%m/%d'
        # convert times
        start = dtparser.parse(gevent.when[0].start_time)
        if start.tzinfo:
            start = start.astimezone(tzlocal())
        else:
            start = start.replace(tzinfo=tzlocal())
        end = dtparser.parse(gevent.when[0].end_time)
        if end.tzinfo:
            end = end.astimezone(tzlocal())
        else:
            end = end.replace(tzinfo=tzlocal())
        # remdict is for file format
        remdict = {'date': start.strftime(remdateformat),
                'time': '', 'dur': '', 'tag': '',
                'summary': '', 'location': '', 'description': ''}
        # remsdict is for output format for uid (rem -s)
        remsdict = {'date': start.strftime(remsdateformat),
                'time': '*', 'dur': '*', 'tag': '*',
                'summary': '*', 'location': '', 'description': ''}
        if start.strftime(remtimeformat) != '00:00':
            remdict['time'] = ' AT ' + start.strftime(remtimeformat)
            remsdict['time'] = (start.time().hour * 60) + \
                    start.time().minute
        # calculate duration
        dur = end - start
        if dur and dur.days != 1:
            durminutes = ((dur.days * 24 * 3600) + dur.seconds) / 60
            duration = '{0[0]}:{0[1]:02}'.format(divmod(durminutes, 60))
            remdict['dur'] = ' DURATION ' + duration
            remsdict['dur'] = durminutes
        # make tags
        cal = None
        cat = None
        for l in gevent.link:
            if l.rel == 'self':
                for piece in l.href.split('/'):
                    if 'default' in piece:
                        cat = 'gigs'
                        continue
                    elif '%40' in piece:
                        cal = piece.replace('%40', '@')
                        continue
        if cal:
            for k, v in caldb['calendars'].items():
                if cal == v:
                    cat = k.lower()
        if cat:
            remdict['tag'] = ' TAG ' + cat
            remsdict['tag'] = cat
        # get transparency
        if gevent.transparency.value == 'TRANSPARENT':
            if remdict['tag'] != '':
                remdict['tag'] += ','
            remdict['tag'] += 'TRANSP=TRANSPARENT'
            if remsdict['tag'] != '*':
                remsdict['tag'] += ',TRANSP=TRANSPARENT'
            else:
                remsdict['tag'] += 'TRANSP=TRANSPARENT'
        # summary, location, description
        remdict['summary'] = gevent.title.text
        if gevent.where[0].value_string:
            remdict['location'] = ' at ' + gevent.where[0].value_string
        if gevent.content.text:
            remdict['description'] = '|\\\n' + \
                    gevent.content.text.replace('\n', '|\\\n')
        self.remline = remstring.format(**remdict)
        # make new uid
        remsdict['body'] = (remdict['summary'] + remdict['location'] + \
                remdict['description']).replace('\n', '').replace('\\', '')
        self.remsstring = remsstring.format(**remsdict)
        self.remuid = hashlib.md5(remsstring.format(**remsdict)).hexdigest()
        # extract filename, linenumber, original uid
        self.filename, self.linenumber, self.origuid = None, None, None
        for ep in gevent.extended_property:
            if ep.name == 'filename':
                self.filename = ep.value
            elif ep.name == 'linenumber':
                self.linenumber = int(ep.value)
            elif ep.name == 'uid':
                self.origuid = ep.value
        # record edit link
        self.link = gevent.GetEditLink().href

    def add_local(self):
        """ add a remevent to the remind file """
        print '# this event was added'
        print self.remline

    def update_local(self):
        """ update a remevent in the remind file """
        print '# this event was changed'
        print '# fileinfo {0} {1}'.format(self.linenumber, self.filename)
        print self.remline

def make_logger(loglevel):
    logger = logging.getLogger(__scriptname__)
    logger.setLevel(loglevel)
    ch = logging.StreamHandler()
    ch.setLevel(loglevel)
    # TODO: alter asctime format
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

if runoptions.verbose:
    logger = make_logger(logging.DEBUG)
else:
    logger = make_logger(logging._levelNames[options['loglevel'].upper()])

def authenticate():
    """attempt to authenticate user"""
    service = gdata.calendar.service.CalendarService()
    service.email = options['user']
    service.password = options['password']
    service.source = __scriptname__
    service.ProgrammaticLogin()
    return service

def get_calendars(service, allcals=False):
    """download user's own calendars from google, or all if allcals=True"""
    if allcals:
        feed = service.GetAllCalendarsFeed()
    else:
        feed = service.GetOwnCalendarsFeed()
    # check/update calendar ids
    caldict = {}
    for c in feed.entry:
        caldict[c.title.text] = c.id.text.rpartition('/')[2].replace('%40', '@')
    caldb['calendars'] = caldict
    caldb.sync()
    return feed

def set_calendars(service, allcals=False):
    """get calendars and (re)set attributes"""
    feed = get_calendars(service, allcals)
    for cal in feed.entry:
        cal.author[0].name.text = 'Mark Knoop'
        cal.timezone = gdata.calendar.Timezone(value='Europe/London')
        cal.where = gdata.calendar.Where(value_string='London')
        calnew = service.UpdateCalendar(calendar=cal)

def new_calendar(service, title, colour='#6e6e41'):
    """make a new calendar, return its id"""
    cal = gdata.calendar.CalendarListEntry()
    cal.title = atom.Title(text=title)
    cal.color = gdata.calendar.Color(value=colour)
    cal.timezone = gdata.calendar.Timezone(value='Europe/London')
    cal.where = gdata.calendar.Where(value_string='London')
    # this sometimes fails, if so try again a few times
    for i in xrange(5):
        try:
            calnew = service.InsertCalendar(new_calendar=cal)
        except RequestError, msg:
            pass
        else:
            break
        raise RequestError, msg
    return calnew.id.text.rpartition('/')[2].replace('%40', '@')

def get_events(service, calid='default', start=None, end=None,
        updatedmin=None, maxresults=1000, query=None):
    """get all events between start and end dates
    start and end should be datetime.date objects"""
    evquery = gdata.calendar.service.CalendarEventQuery(calid, 'private',
            'full', query)
    evquery.max_results = maxresults
    evquery.sortorder = 'ascending'
    evquery.showhidden = 'true' # do I need this to see Android events?
    if updatedmin:
        evquery.updated_min = updatedmin
    if start:
        evquery.start_min = start.strftime(_ddformat)
    if end:
        evquery.start_max = end.strftime(_ddformat)
    evfeed = service.CalendarQuery(evquery)
    return evfeed

def get_all_events(service, updatedmin=None):
    logger.debug(u'Retrieving calendar list.')
    cals = get_calendars(service)
    logger.debug(u'Retrieving event list.')
    events = []
    for calid in caldb['calendars'].values():
        evfeed = get_events(service, calid, updatedmin=updatedmin)
        events += evfeed.entry
    return events

def add_event(service, event):
    """ add a single event """
    gevent = gdata.calendar.CalendarEventEntry()
    gevent.title = atom.Title(text=event.summary)
    if event.location:
        gevent.where.append(gdata.calendar.Where(value_string=event.location))
    if event.description:
        gevent.content = atom.Content(text=event.description)
    gevent.when.append(event.gdatawhen())
    gevent.transparency = gdata.calendar.Transparency()
    gevent.transparency.value = event.transp
    # default settings seem fine for these
    #   * gd:who/gd:attendeeStatus?
    #   * gd:eventStatus
    #   * gd:visibility
    #   * gd:reminder
    # http://code.google.com/apis/gdata/docs/1.0/elements.html#gdEventKind

    # add filename/linenumber/uid as custom properties
    fn = gdata.ExtendedProperty('filename', event.filename)
    ln = gdata.ExtendedProperty('linenumber', event.linenumber)
    uid = gdata.ExtendedProperty('uid', event.uid)
    gevent.extended_property.extend([fn, ln, uid])

    # uri comes from event.categories.value[0]
    #   =>  split categories into different calendars
    cal = 'default'
    if len(event.categories):
        cat = event.categories[0].capitalize()
        if cat in caldb['calendars']:
            cal = caldb['calendars'][cat]
        else:
            # add calendar for new categories
            cal = new_calendar(service, cat)
            caldb['calendars'][cat] = cal
            caldb.sync()
    uri = '/calendar/feeds/{0}/private/full'.format(cal)
    try:
        new_event = service.InsertEvent(gevent, uri)
    except gdata.service.RequestError, msg:
        print msg
        return gevent, False

    logger.debug(u'New event "{0}" added.'.format(event.summary[0:40]))
    return new_event, True

def get_local_calendar():
    """ get local calendar, return dict of (hash, event) """
    # get events from remind
    logger.debug(u'Getting events from remind.')
    # remind options:
    #   -s12    simple output, 12 months
    #   -b2     no times
    #   -l      include fileinfo line
    #   -g      sort
    args = ['/usr/bin/rem', '-s12', '-b2', '-l', '-g']
    # set date as 90 days ago
    args.append(datetime.strftime(datetime.utcnow() - timedelta(days=90),
        _ddformat))
    rem = subprocess.Popen(args, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, close_fds=True).communicate()
    # TODO: close process?
    if rem[1] != '':
        print rem
        logger.error(u'Call to remind failed.')
        return dict()
    remlines = rem[0].lstrip('# fileinfo ').split('\n# fileinfo ')
    events = []
    logger.debug(u'Parsing calendar.')
    for r in remlines:
        events.append(Event(r.decode(_encoding)))

    # check if event is in range
    #evstart = parsedatestring('%s%s%s' % (year, month, day))
    #logging.debug('DTSTART: %s (%s, %s, %s)' % (evstart, year, month, day))
    #if options.fr is not None:
    #    if evstart < options.fr:
    #        continue
    #if options.to is not None:
    #    if evstart > options.to:
    #        continue

    # parse into dictionary of calendar
    localcalendar = dict([(e.uid, e) for e in events])
    return localcalendar

def detect_remote_changes(service, events):
    """ compare events to those in db """
    changed = []
    new = []
    deleted = []
    unchanged, notindb = 0, 0
    logger.info(u'{0} events from Google calendars.'.format(len(events)))
    for e in events:
        rem = Remevent(e)
        if e.event_status.value == 'CANCELED':
            # event has been deleted
            # no point in using rem.remuid since it won't be in db - we can
            # ignore events created and deleted on Google between syncs
            # Google keeps deleted events for a while(?) - this might have been
            # deleted in last sync, so ignore if it's not in the db
            if rem.origuid:
                if rem.origuid in caldb['remotedb']:
                    deleted.append(caldb['remotedb'][rem.origuid])
                    del caldb['remotedb'][rem.origuid]
                else:
                    notindb += 1
        elif rem.origuid is None:
            # event is new
            new.append(rem)
            # TODO: work out how to add fn, ln to events added remotely
            #       perhaps move attendee_status fiddling to Event.add_local()
            #       and do UpdateEvent there
            # events created on Android have
            #  e.who[0].attendee_status = <gdata.calendar.AttendeeStatus object>
            # reset this to None and update event
            if e.who[0].attendee_status is not None:
                e.who[0].attendee_status = None
                e = service.UpdateEvent(e.GetEditLink().href, e)
            # add to remotedb
            caldb['remotedb'][rem.remuid] = (rem.remline, rem.filename,
                    rem.linenumber, e.GetEditLink().href)
        elif rem.remuid != rem.origuid:
            # event has changed
            changed.append(rem)
            # change remotedb
            if rem.origuid in caldb['remotedb']:
                del caldb['remotedb'][rem.origuid]
            else:
                logger.debug(u'Event from Google not in database: {0}.'.format(
                    rem.remline))
            caldb['remotedb'][rem.remuid] = (rem.remline, rem.filename,
                    rem.linenumber, rem.link)
        else:
            # can't detect any change in event
            unchanged += 1
    caldb.sync()
    logger.info(u'{0} new events from Google calendars.'.format(len(new)))
    logger.info(u'{0} changed events from Google calendars.'.format(
            len(changed)))
    logger.info(u'{0} events deleted from Google calendars.'.format(
            len(deleted)))
    logger.info(u'Cannot detect changes in ' +
            u'{0} events from Google calendars.'.format(unchanged))
    logger.debug(u'{0} deleted events from Google calendars '.format(notindb) +
            u'are not in the local database.')
    return new, changed, deleted

def delete_local(remline, fn, ln, link):
    """ delete an event from the remind file """
    print '# this event was deleted'
    print '# fileinfo {0} {1}'.format(ln, fn)
    print remline

def detect_local_changes(localevents):
    """ compare to events in remote db """
    new = set(localevents.keys()) - set(caldb['remotedb'].keys())
    deleted = set(caldb['remotedb'].keys()) - set(localevents.keys())
    logger.info(u'Detected {0} local additions.'.format(len(new)))
    logger.info(u'Detected {0} local deletions.'.format(len(deleted)))
    return list(new), list(deleted)

def delete_all_remote():
    """ delete all remote events """
    logger.debug(u'Logging into Google Calendar.')
    service = authenticate()
    logger.debug(u'Retrieving calendar list.')
    cals = get_calendars(service)
    for calname, calid in caldb['calendars'].items():
        logger.debug(u'Retrieving event list for {0}.'.format(calname))
        evfeed = get_events(service, calid)
        logger.debug(u'Deleting all events from {0}.'.format(calname))
        for e in evfeed.entry:
            service.DeleteEvent(e.GetEditLink().href)

def delete_remote(service, uids):
    """ delete list of events """
    for uid in uids:
        link = caldb['remotedb'][uid][3]
        logger.debug(u'Deleting "{0}" from Google.'.format(uid))
        try:
            service.DeleteEvent(link)
        except gdata.service.RequestError, msg:
            logger.debug(u'...deletion failed: {0}'.format(msg))
        else:
            del caldb['remotedb'][uid]

def add_events(service, uids, events):
    """ add list of events """
    for uid in uids:
        event = events[uid]
        new_gevent, success = add_event(service, event)
        if success:
            # add event details to db
            remevent = Remevent(new_gevent)
            caldb['remotedb'][uid] = (remevent.remline, event.filename,
                    event.linenumber, remevent.link)
        else:
            logger.error(u'Adding event failed.')

def execute():
    """ do sync process """
    # get all events from google
    logger.debug(u'Logging into Google Calendar.')
    service = authenticate()
    updatedmin = None
    if 'lastsync' in caldb.keys() and not runoptions.getall:
        updatedmin = caldb['lastsync']
    remoteevents = get_all_events(service, updatedmin)
    # compare to database
    if 'remotedb' not in caldb:
        caldb['remotedb'] = {}
    new, changed, deleted = detect_remote_changes(service, remoteevents)
    # deal with changes
    for e in new:
        e.add_local()
    for e in changed:
        e.update_local()
    for e in deleted:
        delete_local(*e)
    if len(new) or len(changed) or len(deleted):
        raw_input('Deal with changes, then press return to continue.')
    # get local events
    localevents = get_local_calendar()
    new, deleted = detect_local_changes(localevents)
    # delete remote
    delete_remote(service, deleted)
    caldb.sync()
    # add new events
    add_events(service, new, localevents)
    # set last sync time in config: now() or utcnow()?
    logger.debug(u'Recording sync details.')
    caldb['lastsync'] = datetime.strftime(datetime.utcnow(), _qdtformat)
    caldb.close()

if __name__ == '__main__':
    execute()

