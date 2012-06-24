# Perl Script for Hourly Reporting

## Info
This Perl script takes an .ics calender file as input and produces a .csv file with the hours to be reported on a daily basis. To distinguish between the different projects a file with all PSPs should also be provided as input, as well as a file with the locations of the clients and the month to be processed (see examples below). 


## Required Modules
Data::ICal (http://search.cpan.org/perldoc?Date::ICal)
Date::Manip (http://search.cpan.org/perldoc?Date::Manip)
List::Util qw(sum) (http://search.cpan.org/perldoc?List::Util)
Text::CSV::Encoded (http://search.cpan.org/perldoc?Text::CSV::Encoded) 


## Examples for the command line:
* perl rmeld.pl clientX.ics clientX_psp.csv locations.csv 2012-03

## TODO
* Output a locations file with the TRAVL rows.
* Allow for arbitrary time-periods.