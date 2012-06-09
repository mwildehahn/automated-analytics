====
Python Analyst Library
====

These are a mix of various functions I wrote as an analyst at Eventbrite
(www.eventbrite.com) to automate company reports.

Some are OS dependent (windows_excel_api).

- There are several python libraries that support writing to excel independent
of OS, however these are horrible for maintaining complex excel reports (ie.
with charts, formulas, dynamic tables, formatting etc.) and by horrible I mean
they break all of their functionality. WindowsCom is the best way to easily
write large datasets while maintaining any complex functionality.

The way I set these up to run at Eventbrite was through a windows machine I had
running at the office. Using dropbox, a mix of both pycron* and windows native
task scheduler, this computer would run 24/7 kicking off reports to various
groups & VPs throughout the night. In my opinion, sending out Excel reports is
far superior to static analysis that you could achieve in a variety of other
methods. Being able to provide a lot of complex data to individuals in a format
they are comfortable digging into themselves is extremely beneficial.

Analysts from different departments (marketing & product) leveraged these
functions and processes to automate their reports as well.

* http://www.kalab.com/freeware/pycron/pycron.htm

windows_excel_api.py
----

Example usage::

    import datetime
    from numpy import (
        column_stack,
        vstack,
    )
    from mysqlfunctions import mgdbget
    from windows_excel_api import ExcelApi

    # pull data from a database (MySQL etc.)
    query = """
    SELECT
        count(*) as num_users,
        DATE_FORMAT(created, '%Y-%m-01') as month_created
    FROM Users
    WHERE created > '2012-01-01'
    GROUP BY month_created
    """
    data_set = mgdbget(query)

    >>> data_set
    {'num_users': [341998L, 2775230L, 32307L, 102892233L, 7123158L, 822332L], 'month_created': ['2012-01-01', '2012-02-01', '2012-03-01', '2012-04-01', '2012-05-01', '2012-06-01']}

    # construct a table for output
    headers = ['Month Created', 'Number of Users']
    output_table = column_stack(
        data_set['month_created'],
        data_set['num_users'],
    )

    >>> output_table
    array([['2012-01-01', '744998'],
           ['2012-02-01', '877500'],
           ['2012-03-01', '1041307'],
           ['2012-04-01', '1028923'],
           ['2012-05-01', '950758'],
           ['2012-06-01', '254082']],
          dtype='|S10')

    output_table = vstack((headers, output_table))

    >>> output_table
    array([['Month Created', 'Number of Users'],
           ['2012-01-01', '744998'],
           ['2012-02-01', '877500'],
           ['2012-03-01', '1041307'],
           ['2012-04-01', '1028923'],
           ['2012-05-01', '950758'],
           ['2012-06-01', '254082']],
          dtype='|S15')

    filename = 'mysamplefile.xlsx'

    # way to write to file numerous times
    with ExcelApi() as excel_api:
        excel_api.open_workbook(filename)
        # would have all these already in the file but just showing how to
        # write multiple times
        excel_api.write('Data', 'My automated report', 'A1')
        excel_api.write('Data', str(datetime.datetime.now()), 'A2')
        # write out the data
        excel_api.write('Data', output_table, 'C1')

    # or if you just want to write the output to excel
    xlsxwrite(filename, 'Data', output_table, 'C1')

The way I prefer to structure these reports is have everything in excel
referencing data defined on one sheet. You can have charts, tables, pivot
tables, formulas all pulling in data from this "data" sheet. Then all you
need to do is write a script that handles pulling the data from whatever
sources you have (MySQL, Hadoop, MongoDB etc.), aggregating them together
and doing the processing in python, then writing it out to the "data"
sheet. If you set things up right, you can have fully automated reports
you never need to worry about again.
