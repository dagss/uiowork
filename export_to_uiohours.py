from __future__ import division
from lpod.document import odf_get_document
from time import strptime, mktime, strftime

# TODO: Command-line arguments
TEMPLATE = 'timeregistrering.ods'
YEAR = 2011
OUT = 'timeregistrering-dagss-%d.ods' % YEAR
DATAFILE = 'hours.dat'

def get_week_number(date):
    return int(strftime('%W', date))

def get_first_week_of_month(date):
    return get_week_number(strptime(strftime('%Y-%m-01', date), '%Y-%m-%d'))
    
def parse_time_record(line):
    fields = line.split()
    date = strptime(fields[0], '%Y-%m-%d')
    timerange = fields[1]
    if len(fields) == 3:
        project = fields[2]
    else:
        project = None
        assert len(fields) == 2
    if timerange == 'full':
        timerange = '08:00-15:30'
    start, stop = timerange.split('-')    
    start = strptime(start, '%H:%M')
    stop = strptime(stop, '%H:%M')
    numhours = (mktime(stop) - mktime(start)) / 3600
    return date, numhours, project

def parse_hours_file(filename):
    """
    Load and parse input text file into:
    { date : [ #hours proj. 1, #hours proj., 2, ...] }.

    Returns:

    person, list of projects, time table
    """
    timetable = {}
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
                recidx = projects.index(recproject)
                hourlist[recidx] += rechours
    return person, projects, timetable


def persist_to_ods(template_filename, output_filename, person, projects, timetable):
    document = odf_get_document(template_filename)
    if document.get_type() != 'spreadsheet':
        raise AssertionError()

    # Fetch the tables. The monthly tables are numbers.
    month_tables = [None] * 12
    for tab in document.get_body().get_table_list():
        name = tab.get_name()
        try:
            idx = int(name) # a month?
        except ValueError:
            if name == 'Oversikt':
                oversikt_table = tab
        else:
            month_tables[idx - 1] = tab

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
        week_offset = get_week_number(date) - get_first_week_of_month(date)
        tab = month_tables[date.tm_mon - 1]
        row = 4 + 13 * week_offset
        col = 4 + date.tm_wday
        for p, h in enumerate(hourlist):
            set_value(tab, row + p, col, h)

    document.save(output_filename)
    document = odf_get_document(output_filename)

person, projects, timetable = parse_hours_file(DATAFILE)
persist_to_ods(TEMPLATE, OUT, person, projects, timetable)

