#encoding: utf-8

from __future__ import division
from time import strptime, mktime, strftime
import datetime
import argparse
import sys
from pprint import pprint
from textwrap import dedent
from StringIO import StringIO
import os
import calendar

HOLIDAY_PROJECT = -1
SKIP_PROJECT = -2

def get_week_number(date):
    return int(strftime('%W', date))

def get_first_week_of_month(date):
    return get_week_number(strptime(strftime('%Y-%m-01', date), '%Y-%m-%d'))
    
def parse_time_record(line, prev_date):
    if '#' in line:
        line = line[:line.index('#')]
    fields = line.split()
    if fields[0] == '..':
        if prev_date is None:
            raise ValueError('Cannot use .. as the first date')
        date = prev_date
    else:
        date = strptime(fields[0], '%Y-%m-%d')
    timerange = fields[1]
    if len(fields) == 3:
        project = fields[2]
    else:
        project = None
        assert len(fields) == 2
    if timerange == 'skip':
        project = timerange
        numhours = 0
    elif timerange == 'holiday':
        project = timerange
        numhours = 7.5
    else:
        if timerange == 'full':
            timerange = '08:00-15:30'
        if ';' in timerange:
            timerange, modifier = timerange.split(';')
            assert modifier[-1] == 'm'
            adjustment = int(modifier[:-1]) / 60
        else:
            adjustment = 0
        start, stop = timerange.split('-')    
        start = strptime(start, '%H:%M')
        stop = strptime(stop, '%H:%M')
        numhours = (mktime(stop) - mktime(start)) / 3600 + adjustment
    return date, numhours, project

def parse_hours_file(filename):
    """
    Load and parse input text file into:
    { date : { project : hours worked } }.

    Returns:

    person, time table
    """
    projects = []
    project_names = {}
    timetable = {}
    person = None
    recdate = None
    with file(filename) as f:
        for line in f.readlines():
            line = line.strip()
            if line.startswith('#') or len(line) == 0:
                continue
            elif line.startswith('name:'):
                person = line[len('name:'):].strip()
            elif line.startswith('projects:'):
                for p in [x.strip() for x in line[len('projects:'):].split(',')]:
                    key, desc = p.split(':')
                    project_names[key] = desc
                    projects.append(key)
                default_project = projects[0]
            else:
                recdate, rechours, recproject = parse_time_record(line, recdate)
                if recproject is None:
                    recproject = default_project
                try:
                    hourlist = timetable[recdate]
                except KeyError:
                    hourlist = {}
                    timetable[recdate] = hourlist
                hourlist[recproject] = hourlist.get(recproject, 0) + rechours
    if person is None:
        raise ValueError("person field missing")
    return person, projects, project_names, timetable

def make_report(projects, timetable, last_month=12):
    # Check that all days are accounted for. Rules:
    #  1. Saturday and Sunday off
    #  2. Other days must be registered in one way or another
    dates = timetable.keys()
    dates.sort()
    year = dates[0].tm_year
    ydays_encountered = set(d.tm_yday for d in dates)
    wday = dates[0].tm_wday
    result = True
    owed = 0

    for yday in range(dates[0].tm_yday, dates[-1].tm_yday + 1):
        if (datetime.datetime(year, 1, 1) + datetime.timedelta(days=yday)).month > last_month:
            break
        if wday >= 0 and wday < 5:
            owed += 7.5
            if not yday in ydays_encountered:
                result = False
                print ('WARNING: %s not registered' %
                       strftime('%Y-%m-%d', strptime("%s-%d" % (year, yday), "%Y-%j")))
        wday += 1
        wday %= 7
    summary = dict([(name, 0) for name in projects])
    summary['holiday'] = 0
    for date, work_dict in timetable.iteritems():
        if date.tm_mon > last_month:
            continue
        if date.tm_year != year:
            raise Exception("Two different years in time table: %d and %d" % (year, date.tm_year))
        for project, hourcount in work_dict.iteritems():
            if hourcount > 0: # skip-ed days have hourcount==0
                summary[project] += hourcount
    return result, owed, summary, year

def print_summary(owned, summary):
    def line(desc, value):
        print '%25s %10.1f' % (desc, value)
    def bar():
        print '-' * 36

    s = 0
    keys = summary.keys()
    keys.sort()
    for project in keys:
        hourcount = summary[project]
        line(project, hourcount)
        s += hourcount
    bar()
    line('Sum worked', s)
    line('- 7.5 hours per day', owned)
    bar()
    line('= Balance', s - owned)

def write_tex_report(writer, person, projects, project_names,
                     timetable, month, year, owed, summary):
    month_names = ['januar', 'februar', 'mars', 'april', 'mai', 'juni', 'juli',
                   'august', 'september', 'oktober', 'november', 'desember']
    # Make table
    tabw = StringIO()
    tabw.write(r'\begin{tabular}{ll|%s|l}' % ('c' * len(projects)))
    tabw.write(r'Uke & Dag & %s \\ ' '\n' r'\hline ' '\n' %
               ' & '.join([r'{%s}' % project_names[p] for p in projects]))
    prev_week = None
    for date, hours in sorted(timetable.iteritems()):
        if date.tm_mon != month:
            continue

        week_str = ''
        week = datetime.date(date.tm_year, date.tm_mon, date.tm_mday).isocalendar()[1]
        if prev_week != week:
            # Skip a line
            #if prev_week is not None:
            #    tabw.write(r'\\ ' '\n')
            week_str = str(week)
            prev_week = week

        # Comment if there's anything unusual
        hours_wanted = 7.5 if date.tm_wday <= 4 else 0 
        hours_worked = sum(how_much for what, how_much in hours.iteritems())
        comment = ''
        if 'holiday' in hours:
            comment = 'Off. fridag'
        elif hours_worked < hours_wanted:
            comment = '%.1f t. avspasert' % (hours_wanted - hours_worked)
        elif hours_worked > hours_wanted:
            comment = '%.1f t. oppspart' % (hours_worked - hours_wanted)
        tabw.write('%s & %s & ' % (week_str, date.tm_mday))
        tabw.write(r' & '.join([('%.1f' % hours[p]) if p in hours else '---'
                                   for p in projects]))
        tabw.write(r' & %s \\' % comment)
        tabw.write('\n')
    tabw.write('\end{tabular}')
    table_str = tabw.getvalue()

    # Summary for whole year so far
    sumw = StringIO()
    sumw.write(r'\begin{tabular}{lr}')
    s = 0
    for project in projects:
        sumw.write(r'%s & %.1f \\' % (project_names[project], summary[project]))
        s += summary[project]
    owed -= summary['holiday']
    sumw.write(r'\hline Sum timer ført & %.1f \\' % s)
    sumw.write(r'Antall arbeidsdager $\times$ 7.5 timer & %.1f \\' % owed)
    sumw.write(r'\hline Opptjent avspasering(+)/må jobbe inn(-) & %.1f \\' % (s - owed))
    
    # for date, hours in sorted(timetable.iteritems()):
    #    if date.tm_mon > month:
    #        continue
    #    hours_total = sum(how_much for what, how_much in hours.iteritems())
    #    print >> sys.stderr, date, hours_worked
    #    s += hours_worked
    sumw.write(r'\end{tabular}')
    summary = sumw.getvalue()


    # Load template and substitute
    with file(os.path.split(os.path.abspath(__file__))[0] + '/report_template.tex') as f:
        template = f.read()

    replacements = {'TABLE': table_str,
                    'NAME': person,
                    'MONTH': month_names[month - 1],
                    'YEAR': str(year),
                    'SUMMARY': summary}
    for key, value in replacements.iteritems():
        template = template.replace(key, value)
    writer.write(template)


def status_main(args):
    person, projects, project_names, timetable = parse_hours_file(args.datafile)
    ok, owned, summary, year = make_report(projects, timetable)
    print_summary(owned, summary)

def latex_main(args):
    person, projects, project_names, timetable = parse_hours_file(args.datafile)
    ok, owed, summary, year = make_report(projects, timetable, args.month)
    if args.outfile:
        stream = file(args.outfile, 'w')
    else:
        stream = sys.stdout
    try:
        write_tex_report(stream, person, projects, project_names,
                         timetable, args.month, year,
                         owed, summary)
    finally:
        if args.outfile:
            stream.close()
    return 0


# TODO: Command-line arguments

parser = argparse.ArgumentParser(description='''\
Generate report for worked hours. If outfile is not given, only
report is printed.
''')

subcommands = parser.add_subparsers()
status_parser = subcommands.add_parser('status', help='Sum up hours and show report in terminal')
status_parser.set_defaults(func=status_main)
status_parser.add_argument('datafile', help='Input file where hours worked are listed')


latex_parser = subcommands.add_parser('pdf', help='Make LaTeX file')
latex_parser.set_defaults(func=latex_main)
latex_parser.add_argument('datafile', help='Input file where hours worked are listed')
latex_parser.add_argument('month', type=int)
latex_parser.add_argument('outfile', help='Target output file', nargs='?')

args = parser.parse_args()
sys.exit(args.func(args))
