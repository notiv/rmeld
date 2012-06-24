#!/usr/bin/perl

# example: perl rmeld.pl clientX.ics clientX_psp.csv locations.csv 2012-03

use strict;
use warnings;

use Data::ICal;
use Date::Manip;
use List::Util qw(sum);;
use Text::CSV::Encoded;

# =========== Variables ===============================
my $verbose = 0;

my $dateLow = new Date::Manip::Date;
my $dateHigh = new Date::Manip::Date;
my $err;

my $startTime = new Date::Manip::Date;
my $startTime2 = new Date::Manip::Date;
my $endTime = new Date::Manip::Date;
my $endTime2 = new Date::Manip::Date;
my $delta = new Date::Manip::Delta;

my $str;
my $psp_shortcut;
my $psp_code;
my $psp_code2;
my $this_day;
my $rtext;

my $act_type = "REP1";
my $travel_type = "TRAVL1";
# =====================================================

# =========== Get input arguments =====================
# Get the ics file
my $file = $ARGV[0] or die;

# Get the psp's
$ARGV[1] =~ /\.csv$/ or die;
print "$ARGV[1]\n" if $verbose;
my %psp = &process_psp($ARGV[1]); 	

# Get locations
$ARGV[2] =~ /\.csv$/ or die;
print "$ARGV[2]\n";
my %locations = &process_location($ARGV[2]);

# Get the month from the command line, otherwise use the current date
if( @ARGV > 3 ) {
	$err = $dateHigh->parse($ARGV[3]);
} else {
	$err = $dateHigh->parse("now");
	print "Creating a timesheet for the current month.\n" if $verbose;
}

# =====================================================

# =========== Process date input ======================

# Calculate the first and last days of the month
my $date_string = "last day in ". $dateHigh->printf("%B %Y");

$err = $dateHigh->parse( $date_string );
$dateHigh->set("time",[23,59,59]);

$err = $dateLow->parse( $date_string );
$dateLow->set("d", 1);
$dateLow->set("time",[00,00,00]);

print "First day of the month: " . $dateLow->printf("%Y%m%d") ."\n" if $verbose;
print "Last day of the month: " . $dateHigh->printf("%Y%m%d") ."\n" if $verbose;

my $cal = Data::ICal->new(filename=>$file);
my @temp_events = @{$cal->entries}; # Dereferencing the array reference $cal->entries

# Keep only events of the month, discard all-day events
my @events;
for my $event (@temp_events){
	next unless $event->ical_entry_type eq 'VEVENT'; 				# Only events	
	next unless $event->property('DTSTART')->[0]->value =~ /\dT\d/; # No all-day events
	next unless (($event->property('DTSTART')->[0]->value ge $dateLow->printf("%Y%m%d")) and 
	($event->property('DTSTART')->[0]->value le $dateHigh->printf("%Y%m%d")));		# Only events of the month
	
	# my $tempet = $event->property('DTSTART')->[0]->value;
	# print "$tempet\n";
	
	push @events, $event;
}

my $filename = $ARGV[0];
$filename =~ s/(.*)\.ics/$1/;
print "$filename\n" if $verbose;

my $output_csv = Text::CSV::Encoded->new({ encoding  => "utf8" });
open my $fh_out, ">", "output_" . $filename . ".csv" or die "output.csv: $!";

# For testing purposes
my $input_csv = Text::CSV::Encoded->new({ encoding  => "utf8" });
open my $fh_in, ">", "input_" . $filename . ".csv" or die "input.csv: $!";

my $loc_input_csv = Text::CSV::Encoded->new({ encoding => "utf8" });
open my $fh_loc_in, ">", "loc_input_" . $filename . ".csv" or die "loc_input_csv: $!";

# my $loc_output_csv = Text::CSV::Encoded->new({ encoding => "utf8" });
# open my $fh_loc_out, ">", "loc_output_" . $filename . ".csv" or die "loc_output_csv: $!";

# =====================================================

# =========== Process all events ======================
 
# Hash with timesheet-information
my %timesheet;

# Hash with location information
my %location_sheet;

# Process all events, calculating the time spend on each activity
foreach my $idx1 ( 0 .. $#events ) {
	# Process event
	$startTime->parse_format('%Y%m%dT%H%M%S', $events[$idx1]->property('DTSTART')->[0]->value); # Start time of event
	$endTime->parse_format('%Y%m%dT%H%M%S', $events[$idx1]->property('DTEND')->[0]->value); # End time
	
	$delta = $startTime->calc($endTime);
	
	# Get psp code of event
	if ( $events[$idx1]->property('SUMMARY')->[0]->value =~ /(?<psp_shortcut>\w+):/ ) {
		$psp_code = $psp{$+{psp_shortcut}};	
		} else {
	 	$psp_code = "NA";
	}
		
	# Get current day
	$this_day = $startTime->printf("%Y%m%d");
	
	# Get activity description (text)
	$rtext = $events[$idx1]->property('SUMMARY')->[0]->value;
	
	# For this day accumulate times with the same psp code and description
	$timesheet{$this_day}{$psp_code}{$rtext} += $delta->printf('%.2hhm');
	
	# Get location of event
	my $loc = $events[$idx1]->property('LOCATION')->[0]->value;
	
	# For this day accumulate times with the same psp code and location
	$location_sheet{$this_day}{$psp_code}{$loc} += $delta->printf('%.2hhm');
	
	# Temporary array for writing the input csv (for testing purposes)
	my @rows = ($psp_code, $this_day, $startTime->printf("%Y%m%d%H%M%S"), $endTime->printf("%Y%m%d%H%M%S"), 
		$delta->printf('%.2hhm'), $events[$idx1]->property('SUMMARY')->[0]->value, $idx1);
	
	# Write in csv file
	if ( $input_csv->combine (@rows) ) {
		print $fh_in $input_csv->string, "\n";
	} else {
		print "combine () failed on argument: ", $input_csv->error_input, "\n";
	}
	
	# Temporary array for writing the location input csv (for testing purposes)
	@rows = ($this_day, $loc, $psp_code, $delta->printf('%.2hhm'), $idx1);
	
	# Write in location-input csv file
	if ( $loc_input_csv->combine (@rows) ) {
		print $fh_loc_in $loc_input_csv->string, "\n";
		} else {
			print "combine () failed on argument: ", $loc_input_csv->error_input, "\n";
	}
}

# # Create output location csv (for debugging)
# for my $day (sort keys %location_sheet) {
# 	
# 	# Make an excel-compatible date (for German Version)
# 	my $excel_day = $day;
# 	$excel_day =~ s/(\d{4})(\d{2})(\d{2})/$1.$2.$3/;
# 	
# 	for my $psp ( sort keys %{ $location_sheet{$day} } ) {
# 		for my $loc ( sort keys %{ $location_sheet{$day}{$psp}} ){
# 			
# 			my $hrs = $location_sheet{$day}{$psp}{$loc};
# 					
# 			my @rows = ( $travel_type, $excel_day, $psp, $hrs, $loc);
# 			print "@rows\n";
# 			
# 			# Write in csv file
# 			if ( $loc_output_csv->combine (@rows) ) {
# 				print $fh_loc_out $loc_output_csv->string, "\n";
# 			} else {
# 				print "combine () failed on argument: ", $loc_output_csv->error_input, "\n";
# 			}
# 		}
# 	}
# }


# Create output csv
for my $day (sort keys %timesheet) {
	
	# Make an excel-compatible date (for German Version)
	my $excel_day = $day;
	$excel_day =~ s/(\d{4})(\d{2})(\d{2})/$1.$2.$3/;
	
	for my $psp ( sort keys %{ $timesheet{$day} } ) {
		for my $rtext ( sort keys %{ $timesheet{$day}{$psp}} ){
			
			my $hrs = $timesheet{$day}{$psp}{$rtext};
					
			my @rows = ( $act_type, $excel_day, $psp, $hrs, $rtext);
			print "@rows\n";
			
			# Write in csv file
			if ( $output_csv->combine (@rows) ) {
				print $fh_out $output_csv->string, "\n";
			} else {
				print "combine () failed on argument: ", $output_csv->error_input, "\n";
			}
		}
	}

}

 


sub process_psp{
	my $file = $_[0];
	
	my $csv = Text::CSV::Encoded->new({ encoding  => "utf8" });
	my %psp;
	
	open (CSV, "<", $file) or die $!;
	
	while (<CSV>) {
		if ($csv->parse($_)) {
			my @columns = $csv->fields();
			
			if ($columns[0] ne ""){
				$psp{$columns[0]}=$columns[1];
				print "$columns[0] \t -> \t $columns[1]\n" if $verbose; 
			}
		} else {
			my $err = $csv->error_input;
			print "Failed to parse line: $err";
		}
	}
	%psp;
}

sub process_location{
	my $file = $_[0];
	# print "$file\n";
	
	my $csv = Text::CSV::Encoded->new({ encoding  => "utf8" });
	my %locations;
	
	open (CSV, "<", $file) or die $!;
	
	while (<CSV>) {
		if ($csv->parse($_)) {
			my @columns = $csv->fields();
			
			if ($columns[0] ne ""){
				$locations{$columns[0]}=$columns[1];
				print "$columns[0] \t -> \t $columns[1]\n" if $verbose; 
			}
		} else {
			my $err = $csv->error_input;
			print "Failed to parse line: $err";
		}
	}
	%locations;
}
