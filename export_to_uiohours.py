from __future__ import division
from lpod.document import odf_get_document
from time import strptime, mktime, strftime
import argparse
import sys

HOLIDAY_PROJECT = -1
SKIP_PROJECT = -2

def get_week_number(date):
    return int(strftime('%W', date))

def get_first_week_of_month(date):
    return get_week_number(strptime(strftime('%Y-%m-01', date), '%Y-%m-%d'))
    
def parse_time_record(line):
    if '#' in line:
        line = line[:line.index('#')]
    fields = line.split()
    date = strptime(fields[0], '%Y-%m-%d')
    timerange = fields[1]
    if len(fields) == 3:
        project = fields[2]
    else:
        project = None
        assert len(fields) == 2
    if timerange in ('skip', 'holiday'):
        project = timerange
        numhours = 0
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
    { date : [ #hours proj. 1, #hours proj., 2, ...] }.

    Returns:

    person, list of projects, time table
    """
    timetable = {}
    person = None
    with file(filename) as f:
        for line in f.readlines():
            line = line.strip()
            if line.startswith('#') or len(line) == 0:
                continue
            elif line.startswith('name:'):
                person = line[len('name:'):].strip()
            elif line.startswith('projects:'):
                projects = [x.strip() for x in line[len('projects:'):].split(',')]
                default_project = projects[0]
            else:
                recdate, rechours, recproject = parse_time_record(line)
                if recproject is None:
                    recproject = default_project
                try:
                    hourlist = timetable[recdate]
                except:
                    hourlist = [0] * len(projects)
                    timetable[recdate] = hourlist
                if recproject == 'holiday':
                    recidx = HOLIDAY_PROJECT
                elif recproject == 'skip':
                    recidx = SKIP_PROJECT
                else:
                    recidx = projects.index(recproject)
                hourlist[recidx] += rechours
    if person is None:
        raise ValueError("person field missing")
    return person, projects, timetable

def sanity_check(timetable):
    # Check that all days are accounted for. Rules:
    #  1. Saturday and Sunday off
    #  2. Other days must be registered in one way or another
    dates = timetable.keys()
    dates.sort()
    year = dates[0].tm_year
    ydays_encountered = set(d.tm_yday for d in dates)
    wday = dates[0].tm_wday
    result = True
    for yday in range(dates[0].tm_yday, dates[-1].tm_yday + 1):
        if wday >= 0 and wday < 5:
            if not yday in ydays_encountered:
                result = False
                print ('WARNING: %s not registered' %
                       strftime('%Y-%m-%d', strptime("%s-%d" % (year, yday), "%Y-%j")))
        wday += 1
        wday %= 7
    return result
        

def persist_to_ods(template_filename, output_filename, person, projects, timetable):
    document = odf_get_document(template_filename)
    if document.get_type() != 'spreadsheet':
        raise AssertionError()

    # Fetch the tables. The monthly tables are numbers.
    month_tables = [None] * 12
    first_week_list = [None] * 12
    for tab in document.get_body().get_table_list():
        name = tab.get_name()
        try:
            idx = int(name) # a month?
        except ValueError:
            if name == 'Oversikt':
                oversikt_table = tab
        else:
            month_tables[idx - 1] = tab
            first_week_list[idx - 1] = int(tab.get_row(2).get_cell(2).get_value())

    def set_value(tab, row, col, value):
        # For some reason, direct set_value on table does not work
        r = tab.get_row(row)
        c = r.get_cell(col)
        c.set_value(value)
        r.set_cell(col, c)
        tab.set_row(row, r)

    # Register meta-information
    set_value(oversikt_table, 6, 4, person)
    set_value(oversikt_table, 10, 4, projects[0])
    for idx, projname in enumerate(projects[1:]):
        set_value(oversikt_table, 17 + idx, 2, 150300)
        set_value(oversikt_table, 17 + idx, 3, 890000)
        set_value(oversikt_table, 17 + idx, 4, projname)

    # Register times
    for date, hourlist in timetable.iteritems():
        month_idx = date.tm_mon - 1
        week = get_week_number(date)
        week_offset = week - first_week_list[month_idx]
        tab = month_tables[month_idx]
        week_present_in_doc = int(tab.get_row(2 + 13 * week_offset).get_cell(2).get_value())
        if week_present_in_doc != week:
            # Sanity check
            raise AssertionError("%d != %d" % (week_present_in_doc, week))
        row = 4 + 13 * week_offset
        col = 4 + date.tm_wday
        for projidx, h in enumerate(hourlist):
            if projidx in (HOLIDAY_PROJECT, SKIP_PROJECT):
                continue # holiday
            set_value(tab, row + projidx, col, h)

    document.save(output_filename)
    document = odf_get_document(output_filename)

def main(args):
    if args.datafile == args.template:
        raise Exception("Trying to overwrite template file...")
    person, projects, timetable = parse_hours_file(args.datafile)
    if not sanity_check(timetable):
        sys.stderr.write('ERROR: Fix errors, then I will move on\n')
        return 1
    else:
        persist_to_ods(args.template, args.outfile, person, projects, timetable)
        return 0


# TODO: Command-line arguments

parser = argparse.ArgumentParser(description='Precomputation')

parser.add_argument('-t', '--template', default='timeregistrering.ods',
                    help='Template ODS file')

parser.add_argument('datafile', help='Input file where hours worked are listed')
parser.add_argument('outfile', help='Target output file')

sys.exit(main(parser.parse_args()))
