#!/usr/bin/env perl

use strict;
use warnings;
use experimental qw(smartmatch switch); # for ~~ and given ... when

use CPAN;
use CPAN::Shell;

my %required_modules = (
	"Archive::Zip"                 => undef,
	"Carp"                         => undef,
	"criticism"                    => ['harsh'], # Checks code style according to Perl Best Practices: 
							                     # 5 - gentle (the weakest), stern, harsh, cruel , 1 - brutal (the strongest)
	"Crypt::Checksum"              => ['crc32_file_hex'],
	"Data::Dumper"                 => undef,
	"DateTime"                     => undef,
	"DateTime::Format::Strptime"   => undef,
	"Excel::Writer::XLSX::Utility" => undef,
	"File::Copy"                   => undef,
	"File::Find"                   => undef,
	"File::Remove"                 => undef,
	"File::Spec"                   => undef,
	"IO::Handle"                   => undef,
	"JIRA::REST"                   => undef,
	"JSON"                         => undef,
	"Statistics::Descriptive"      => undef,
	"String::ProgressBar"          => undef,
	"Win32::OLE"                   => undef,
	"WWW::Github::Files"           => undef,
);

print "Checking installed modules..\n";
check_and_install_required_modules(\%required_modules);

print "\nImport required modules..\n";
import_required_modules(\%required_modules);

STDOUT->autoflush(1);           # autoflush for STDOUT

open my $input_file_handle, "<", "settings.json" or croak $!;
my @lines = <$input_file_handle>;
close $input_file_handle;

my $json     = JSON->new->allow_nonref;
my $configuration = $json->decode(join('', @lines));

check_for_updates($configuration);
check_mandatory_configuration($configuration);
add_additional_configuration($configuration);

my $jira_settings = $configuration->{jira};
my $jira     = connect_to_jira($jira_settings->{url}, $jira_settings->{login}, $jira_settings->{password});
my $data     = load_query_data($jira, $configuration);

create_report($configuration, $data);

sub create_report {
	my $conf_obj   = shift;
	my $query_data = shift;
	my $settings = $conf_obj->{settings};
	my $template_file_name = $settings->{template_file};
	my $report_file_name   = $settings->{report_file};
	my $copy_template_to_report = $settings->{copy_template_to_report};
	
	$report_file_name   = convert_file_name_to_absolute($report_file_name);
	$template_file_name = convert_file_name_to_absolute($template_file_name);
	
	print "\nCreating report $report_file_name: ";
	
	my $excel = create_excel_object();
	
	create_report_file_from_template($excel, $template_file_name, $report_file_name, $copy_template_to_report);
	replace_keys_to_data($excel, $report_file_name, $query_data);
	
	$excel->Quit();
	
	print "Done\n";
	
	return;
}

sub replace_keys_to_data {
	my $excel            = shift;
	my $report_file_name = shift;
	my $query_data       = shift;

	my $workbook = $excel->Workbooks->Open($report_file_name);
	my $worksheets_count = $workbook->Worksheets->Count;

	foreach my $sheet_number (1..$worksheets_count)
	{
		my $worksheet = $workbook->Worksheets($sheet_number);
		$worksheet->Activate();
		
		my $last_row = $worksheet->UsedRange->Rows->Count;
		my $last_column = $worksheet->UsedRange->Columns->Count;
		
		foreach my $row (0..$last_row) {
			foreach my $column (0..$last_column) {
				my $cell_address = xl_rowcol_to_cell($row, $column);
				my $cell_value = $worksheet->Range($cell_address)->Value;
				
				if(defined $cell_value and defined $query_data->{uc("$cell_value")}) {
					$worksheet->Range($cell_address)->{Value} = "" . $query_data->{uc("$cell_value")};
				}
			}
		}
		
		$worksheet->Range("A1")->Select();
	}
	
	$workbook->Worksheets(1)->Activate();
	
	$workbook->Save();
	$workbook->Close();
	
	return;
}

sub create_excel_object {
	my $automation_object_class = "Excel.Application";
	my $excel = Win32::OLE->new($automation_object_class, 'Quit')          # run new instance
		or confess("Can't run new instance of Excel automation object");   # or stop script with backtrace
	$excel->{DisplayAlerts} = 0;
	
	return $excel;
}

sub create_report_file_from_template {
	my $excel = shift;
	my $template_file_name = shift;
	my $report_file_name = shift;
	my $copy_template_to_report = shift;
	
	if(defined $copy_template_to_report and $copy_template_to_report) {
		my $template_workbook = $excel->Workbooks->Open($template_file_name);
		my $report_workbook   = $excel->Workbooks->Open($report_file_name);
		my $worksheets_count = $template_workbook->Worksheets->Count;
		
		foreach my $sheet_number (1..$worksheets_count) {
			my $template_worksheet = $template_workbook->Worksheets($sheet_number);
			my $sheet_name = $template_worksheet->{Name};
			my $report_worksheet = $report_workbook->Worksheets("$sheet_name");
			my $last_column = $report_worksheet->UsedRange->Columns->Count;
			
			my $range_for_copy_from_template = xl_range_formula("$sheet_name", 
				0, $template_worksheet->UsedRange->Rows->Count,     # start and end rows
				0, $template_worksheet->UsedRange->Columns->Count); # start and end columns
			
			$template_worksheet->Range($range_for_copy_from_template)->Copy();
			
			my $start_cell_for_insert_in_report = xl_rowcol_to_cell(0, $last_column);
			$report_worksheet->Range($start_cell_for_insert_in_report)->PasteSpecial();
			
			$template_workbook->Save();
			$template_workbook->Close();
			
			$report_workbook->Save();
			$report_workbook->Close();
		}
	} else {
		copy("$template_file_name", "$report_file_name");
	}
	
	return;
}

sub convert_file_name_to_absolute {
	my $file_name = shift;
	
	if(not File::Spec->file_name_is_absolute($file_name)) {
		$file_name = File::Spec->rel2abs($file_name);
	}
	
	return $file_name;
}

sub convert_file_name_to_relative {
	my $file_name = shift;
	
	if(File::Spec->file_name_is_absolute($file_name)) {
		$file_name = File::Spec->abs2rel($file_name);
	}
	
	return $file_name;
}

sub load_query_data {
	my $jira_obj = shift;
	my $conf_obj = shift;
	
	my %result = ();
	
	foreach my $query (@{$conf_obj->{queries}}) {
		next if not defined $query->{template_key} or not $query->{template_key};
		
		print "\nGetting data for $query->{template_key}...\n";
		
		my ($issues, $subtasks) = search_issues($jira, $query);
		$result{uc("$query->{template_key}")} = "" . ($#{$issues} + 1);
		$result{uc("report_date")} = "" . today();
		$result{uc("work_week_number")} = "" . get_work_week_number();
		
		add_period_data($query, \%result);
		add_priority_data($query, $issues, \%result);
		add_status_data($conf_obj, $query, $issues, \%result);
		add_time_tracking_data($conf_obj, $query, $issues, $subtasks, \%result)
	}
	
	my $save_keys_into_file = $conf_obj->{settings}->{save_keys_into_file};
	
	if(defined $save_keys_into_file and $save_keys_into_file) {
		my @keys = sort keys %result;
		open my $output_file_handle, ">", "available_keys.txt" or croak $!;
		print $output_file_handle "$_=$result{$_}\n" foreach @keys;
		close $output_file_handle;
	}
	
	return \%result;
}

sub today {
	my $formatter  = DateTime::Format::Strptime->new(pattern => '%F');
	my $dt = DateTime->now();
	
	$dt->set_formatter($formatter);
	
	return $dt;
}

sub get_work_week_number {
	my $formatter  = DateTime::Format::Strptime->new(pattern => '%j');
	my $dt = DateTime->now();
	$dt->set_formatter($formatter);
	my $day_number = "" . $dt;        # The day number in the year (1-366).
	
	$formatter  = DateTime::Format::Strptime->new(pattern => '%u');
	$dt = DateTime->now();
	$dt->set_formatter($formatter);
	$dt->truncate( to => 'year');
	my $delta = "" . $dt - 1;         # how many days between 1 Jan and Monday. 
                                      # If 1 Jan is Monday (number of Monday is 1) then delta is 0, 
									  # If 1 Jan is Tuesday (number of Tuesday is 2) then delta is 1, etc..
	
	my $week_number =  int(($day_number + $delta - 1) / 7) + 1;
	
	return $week_number;
}

sub search_issues {
	my $jira_obj = shift;
	my $query = shift;
	my @issues = ();
	my @subtasks = ();
	my $search = undef;
	my $startAt = 0;
	my $very_big_value = 100000;
	my $progress_bar = undef;
	my $total_issues = undef;
	
	do {
		$search = $jira_obj->GET('/search', {
			'jql' => $query->{query},
			'expand' => "changelog",
			"startAt" => "$startAt"
		});
		
		if(not defined $progress_bar and $search->{total} > 0 and $search->{total} > $search->{maxResults}) {
			$total_issues = $search->{total};
			
			$progress_bar = String::ProgressBar->new( 
				max          => $total_issues, 
				length       => 60,
				print_return => 1
			);
		}
		
		push @issues, @{$search->{issues}};
		my $number_of_subtasks = add_subtasks($jira_obj, $search->{issues}, \@subtasks);
		
		if(defined $number_of_subtasks and $number_of_subtasks) {
			$total_issues = $search->{total} if not defined $total_issues;
			$total_issues += $number_of_subtasks;
			
			if($total_issues > $search->{maxResults}) {
				$progress_bar = String::ProgressBar->new( 
					max          => $total_issues, 
					length       => 60,
					print_return => 1
				);
			}
		}
		
		$startAt += $search->{maxResults};
		
		if(defined $progress_bar and $search->{total} > 0) {
			$progress_bar->update($#issues + 1);
			$progress_bar->write;
		}
		
	} while ($search->{startAt} + $#{$search->{issues}} + 1 < $search->{total});
	
	$total_issues = $search->{total} if not defined $total_issues;
	print "Wrong number of issues!\n" . 
		"Expected: $search->{total}\n" . 
		"Actual: "   . ($#issues + 1) . 
		"\n" if $search->{total} != ($#issues + 1);
	
	#print Dumper(\@issues);
	
	return \@issues, \@subtasks;
}

sub add_subtasks {
	my $jira_obj      = shift;
	my $parent_issues = shift;
	my $issues        = shift;
	my $number_of_subtasks = 0;
	
	my @subtasks_keys = ();
	my @issues_with_subtasks = grep { defined $_->{fields}->{subtasks} and $#{$_->{fields}->{subtasks}} >= 0 } @{$parent_issues};
	
	foreach my $issue (@issues_with_subtasks) {
		foreach my $subtask (@{$issue->{fields}->{subtasks}}) {
			push @subtasks_keys, "key=$subtask->{key}";
		}
	}
	
	if($#subtasks_keys >= 0) {
		my $query   = join(" or ", @subtasks_keys);
		my $startAt = 0;
		my $search  = undef;
		
		do {
			$search = $jira_obj->GET('/search', {
				'jql' => "$query",
				'expand' => "changelog",
				"startAt" => "$startAt"
			});
			
			push @{$issues}, @{$search->{issues}};
			
			$number_of_subtasks += $#{$search->{issues}} + 1;
			$startAt += $search->{maxResults};
		} while ($search->{startAt} + $#{$search->{issues}} + 1 < $search->{total});
	}
	
	return $number_of_subtasks;
}

sub add_time_tracking_data {
	my $conf_obj = shift;
	my $query    = shift;
	my $issues   = shift;
	my $subtasks = shift;
	my $result   = shift;
	my $key      = undef;
	
	return if not defined $query->{get_estimation_time};
	
	my @fields   = @{$query->{get_estimation_time}};
	$_ = lc foreach @fields;                          # make all values lower case
	
	foreach my $field (@fields) {
		my @estimated_time = get_time($conf_obj, $field, $issues, $subtasks);
		$key = replace_spaces_by_emphasis(uc("$query->{template_key}_$field"));
		add_part_of_data($conf_obj, $conf_obj->{fields}->{$field}, $result, $key, @estimated_time);
		
		if(defined $query->{order_by_priority}) {
			my @priorities = @{$query->{order_by_priority}};
			
			foreach my $priority (@priorities) {
				my $priority_regexp = make_regexp_from_string($priority);
				my @estimated_time_for_priority = grep {$_->{priority} =~ /^$priority_regexp$/ix} @estimated_time;
				$key = replace_spaces_by_emphasis(uc("$query->{template_key}_${field}_$priority"));
				add_part_of_data($conf_obj, $conf_obj->{fields}->{$field}, $result, $key, @estimated_time_for_priority);
			}
		}
	}
		
	return;
}

sub get_time {
	my ($conf_obj, $field, $issues, $subtasks) = @_;
	my @times = ();
	
	foreach my $issue (@{$issues}) {
		my @data_for_times = process_history_items($issue, $conf_obj->{fields}->{$field});
		my $time_object = pop @data_for_times;
		my $time = $time_object->{to};
		
		if(defined $issue->{fields}->{subtasks} and $#{$issue->{fields}->{subtasks}} >= 0) {
			my @subtasks_keys = get_subtasks_keys($issue);
			my @subtasks = grep { $_->{key} ~~ @subtasks_keys } @{$subtasks};
			
			foreach my $subtask (@subtasks) {
				@data_for_times = process_history_items($subtask, $conf_obj->{fields}->{$field});
				my $subtask_time_object = pop @data_for_times;
				$time += $subtask_time_object->{to};
			}
		}
		
		$time = convert_seconds_to_hours($time);
		
		push @times, {
			"$conf_obj->{fields}->{$field}" => "$time",
			priority => "$time_object->{priority}",
			issue => "$time_object->{issue}"
		};
		
	}
	
	return @times;
}

sub convert_seconds_to_hours {
	my $value = shift;
	my $seconds_in_hour = 3600;
	
	return $value / $seconds_in_hour;
}

sub get_subtasks_keys {
	my @issues = @_;
	my @subtasks_keys = ();
	
	foreach my $issue (@issues) {
		next if not defined $issue->{fields}->{subtasks};
		next if $#{$issue->{fields}->{subtasks}} < 0;
		
		foreach my $subtask (@{$issue->{fields}->{subtasks}}) {
			push @subtasks_keys, $subtask->{key};
		}
	}
	
	return @subtasks_keys;
}

sub add_status_data {
	my $conf_obj = shift;
	my $query    = shift;
	my $issues   = shift;
	my $result   = shift;

	return if not defined $query->{get_time_for_statuses};
	
	my @times = ();
	my @statuses = @{$query->{get_time_for_statuses}};
	$_ = lc($_) foreach @statuses; # it's not neccessary but let it be
	
	foreach my $issue (@{$issues}) {
		my @data_for_times = process_history_items($issue, $conf_obj->{fields}->{status});
		my @times_for_issue = get_status_times(@data_for_times);
		push @times, @times_for_issue if $#times_for_issue >= 0;
	}
	
	foreach my $status (@statuses) {
		my $status_regexp = make_regexp_from_string($status);
		my @times_for_status = grep { $_->{status} =~ /^$status_regexp$/ix } @times;
		my $key = replace_spaces_by_emphasis(uc("$query->{template_key}" . "_${status}"));
		
		add_part_of_data($conf_obj, $conf_obj->{fields}->{status}, $result, $key, @times_for_status);
		
		if(defined $query->{order_by_priority}) {
			my @priorities = @{$query->{order_by_priority}};
			
			foreach my $priority (@priorities) {
				my $priority_regexp = make_regexp_from_string($priority);
				
				@times_for_status   = grep { 
					$_->{status} =~ /^$status_regexp$/ix 
					and $_->{priority} =~ /^$priority_regexp$/ix 
				} @times;
				
				$key = replace_spaces_by_emphasis(uc("$query->{template_key}" . "_${status}_${priority}"));

				add_part_of_data($conf_obj, $conf_obj->{fields}->{status}, $result, $key, @times_for_status);
			}
		}
	}
	
	return;
}

sub make_regexp_from_string {
	my $string = shift;
	my $regexp = $string;
	
	$regexp =~ s/\s+/\\s+/ixg;
	
	return $regexp;
}

sub add_part_of_data {
	my ($conf_obj, $field, $result, $key, @times) = @_;
	my @data_for_statistic = ();
	
	given($field) {
		when(/^$conf_obj->{fields}->{status}$/ix) {
			@data_for_statistic = get_values_from_hash("period", @times);
			add_statistic_data($conf_obj, $result, "$key", @data_for_statistic);
		}
		when(/^$conf_obj->{fields}->{estimated}$/ix) {
			@data_for_statistic = get_values_from_hash("$conf_obj->{fields}->{estimated}", @times);
			add_statistic_data($conf_obj, $result, "$key", @data_for_statistic);
			$result->{uc("$key")} = $result->{uc("${key}_sum")};
		}
		when(/^$conf_obj->{fields}->{logged}$/ix) {
			@data_for_statistic = get_values_from_hash("$conf_obj->{fields}->{logged}", @times);
			add_statistic_data($conf_obj, $result, "$key", @data_for_statistic);
			$result->{uc("$key")} = $result->{uc("${key}_sum")};
		}
		when(/^$conf_obj->{fields}->{remaining}$/ix) {
			@data_for_statistic = get_values_from_hash("$conf_obj->{fields}->{remaining}", @times);
			add_statistic_data($conf_obj, $result, "$key", @data_for_statistic);
			$result->{uc("$key")} = $result->{uc("${key}_sum")};
		}
		default {
			confess("Unexpected field for adding data!");
		}
	}
		
	my @issues_keys = get_issue_keys(@times);
	$result->{uc("${key}_JIRA_LINK")} = get_jira_link(@issues_keys);
	
	return;
}

sub get_values_from_hash {
	my ($field, @data) = @_;
	
	my @values = ();
	push @values, $_->{"$field"} foreach @data;
	
	return @values;
}

sub get_issue_keys {
	my @times_for_status = @_;
	
	my @keys = ();
	push @keys, { key => $_->{issue}} foreach @times_for_status;
	
	return @keys;
}

sub add_statistic_data {
	my ($conf_obj, $result, $key, @values) = @_;
	
	my $stat = Statistics::Descriptive::Full->new();
	$stat->add_data(@values);
	
	my $min     = $stat->min();
	my $max     = $stat->max();
	my $average = $stat->mean();
	my $median  = $stat->median();
	my $stddev  = $stat->standard_deviation();
	my $sum     = $stat->sum();
	
	$result->{uc("$key" . "_min")}     = replace_undefined_value_to_default($conf_obj, $min);
	$result->{uc("$key" . "_max")}     = replace_undefined_value_to_default($conf_obj, $max);
	$result->{uc("$key" . "_average")} = replace_undefined_value_to_default($conf_obj, $average);
	$result->{uc("$key" . "_median")}  = replace_undefined_value_to_default($conf_obj, $median);
	$result->{uc("$key" . "_stddev")}  = replace_undefined_value_to_default($conf_obj, $stddev);
	$result->{uc("$key" . "_sum")}     = replace_undefined_value_to_default($conf_obj, $sum);
	
	return;
}

sub replace_undefined_value_to_default {
	my $conf_obj = shift;
	my $value    = shift;
	my $default_from_settings = $conf_obj->{settings}->{default_value_for_data};
	my $default_value = (defined $default_from_settings) ? "$default_from_settings" : "n/a";
	
	return (defined $value) ? "$value" : "$default_value";
}

sub replace_spaces_by_emphasis {
	my $value  = shift;
	my $result = $value;
	
	$result =~ s/\s/_/igx;
	$result =~ s/_+/_/igx;
	
	return $result;
}

sub add_priority_data {
	my $query  = shift;
	my $issues = shift;
	my $result = shift;
	
	return if not defined $query->{order_by_priority};

	my @priorities = @{$query->{order_by_priority}};
	
	foreach my $priority (@priorities) {
		my @issues_with_priority = grep {$_->{fields}->{priority}->{name} =~ /^$priority$/ix} @{$issues};
		$result->{uc("$query->{template_key}" . "_$priority")} = "" . $#issues_with_priority + 1;
		$result->{uc("$query->{template_key}" . "_${priority}_JIRA_LINK")} = get_jira_link(@issues_with_priority);
	}
	
	return;
}

sub add_period_data {
	my $query  = shift;
	my $result = shift;
	
	if(defined $query->{period} and $query->{period}) {
		$result->{uc("$query->{template_key}" . "_PERIOD_NAME")} = "" . get_period_name($query->{period});
	}
	
	return;
}

sub get_status_times {
	my @data_for_times = @_;
	
	my @result_times  = ();
	my $previous_date = undef;
	my $current_date  = undef;
	
	foreach my $data_item (@data_for_times) {
		if($data_item->{from} =~ /^open$/ix ) {
			$previous_date = $data_item->{created};
		}
		
		print "Previous date is undefined!\n" if not defined $previous_date; 
		
		$current_date = $data_item->{date};
		
		push @result_times, {
			status   => "$data_item->{from}",
			period   => "" . get_period($previous_date, $current_date), 
			priority => "$data_item->{priority}",
			issue    => $data_item->{issue}
		};
		
		$previous_date = $current_date;
	}
	
	return @result_times;
}

sub process_history_items {
	my $issue = shift;
	my $field = shift;
	
	my $strp = DateTime::Format::Strptime->new(
		pattern  => '%Y-%m-%dT%T.%3N%z',
		zone_map => { UTC => "+0300"}
	);
	
	my @data_for_times    = ();
	my @all_history_items = @{$issue->{changelog}->{histories}};
	$field = make_regexp_from_string($field);
	
	foreach my $history_item (@all_history_items) {
		if(has_changes_in_field($field, $history_item->{items}, )) {
			my @status_items = grep { $_->{field} =~ /^$field$/ix } @{$history_item->{items}}; # filter only status changes
				
			print "More then one status item!\n" if $#status_items;                         # it shouldn't happen but anyway
			
			my $item = shift @status_items;                                                 # process only first item
			
			push @data_for_times, { 
				created  => $strp->parse_datetime($issue->{fields}->{created}),
				date     => $strp->parse_datetime($history_item->{created}),
				from     => $item->{fromString}, 
				to       => $item->{toString}, 
				priority => $issue->{fields}->{priority}->{name},
				issue    => $issue->{key}
			};
		} 
	}
	
	@data_for_times = sort { DateTime->compare($a->{date}, $b->{date}) } @data_for_times;
	
	return @data_for_times;
}


sub get_period { # in hours
	my $start_date  = shift;
	my $end_date    = shift;
	
	my $seconds_in_minute = 60;
	my $minutes_in_hour   = 60;
	my $hours_in_day      = 24;
	my $period            = 0; 
	
	my $duration = $end_date - $start_date;
	my %deltas   = $duration->deltas();
		
	# convert period to seconds
	$period += $deltas{seconds} if $deltas{seconds};
	$period += ($deltas{minutes} * $seconds_in_minute) if $deltas{minutes};
	$period += ($deltas{hours} * $minutes_in_hour * $seconds_in_minute) if $deltas{hours};
	$period += ($deltas{days} * $minutes_in_hour * $seconds_in_minute * $hours_in_day) if $deltas{days};
	
	return $period / $seconds_in_minute / $minutes_in_hour; # convert seconds to hours
}

sub has_changes_in_field {
	my $field = shift;
	my $items = shift;
	
	my $field_regexp = make_regexp_from_string($field);
	
	my @status_items = grep { $_->{field} =~ /^$field_regexp$/ix } @{$items};
	
	return $#status_items >= 0;
}

sub get_period_name {
	my $period     = shift;
	my $month_name = "";
	my $formatter  = DateTime::Format::Strptime->new(pattern => '%B');
	my $dt         = undef;
	
	if($period =~ /^month_(\d{1,2})$/ix) {
		my $month_for_subtract = $1;
		
		$dt = DateTime->now();
		$dt->set_formatter($formatter);
		
		if($month_for_subtract) {
			$dt->subtract(months => $month_for_subtract);
		}
		
		$month_name = $dt;
	} else {
		$month_name = "last week";
	}
	
	return $month_name;
}

sub get_jira_link {
	my @issues    = @_;
	my @jira_keys = ();
	
	push @jira_keys, $_->{key} foreach @issues;
	@jira_keys = uniq(@jira_keys);
	
	return "" if $#jira_keys < 0;
	
	return "https://jira.returnonintelligence.com/issues/?jql=key in (" . join(", ", @jira_keys) . ")";
}

sub uniq {
	my @data    = @_;
	my @results = ();
	my %seen    = ();
	
	foreach my $value (@data) {
		if(!$seen{$value}) {
			push @results, $value;
			$seen{$value} = 1;
		}
	}
	
	return @results;
}

sub connect_to_jira {
	my ($url, $login, $password) = @_;
	
	return JIRA::REST->new($url, $login, $password);
}

sub check_mandatory_configuration {
	my $conf_obj = shift;
	my %mandatory_conf = (
		jira     => [qw(login password url)],
		settings => [qw(template_file report_file)]
	);
	my $config_has_missed_parameters = 0;
	
	foreach my $section (keys %mandatory_conf) {
		if(not defined $conf_obj->{"$section"}) {
			print "Section '$section' is missed in configuration\n";
			$config_has_missed_parameters = 1;
		} elsif(not $conf_obj->{"$section"}) {
			print "Section '$section' is empty or zero in configuration\n";
			$config_has_missed_parameters = 1;
		}
		
		foreach my $parameter (@{$mandatory_conf{"$section"}}) {
			my $parameter_in_config = $conf_obj->{"$section"}->{"$parameter"};
			
			if(not defined $parameter_in_config) {
				print "Parameter '$parameter' is missed in section '$section'\n";
				$config_has_missed_parameters = 1;
			} elsif (not $parameter_in_config) {
				print "Parameter '$parameter' is empty or zero in section '$section'\n";
				$config_has_missed_parameters = 1;
			}
		}
	}
	
	exit 1 if $config_has_missed_parameters;
	
	return;
}

sub add_additional_configuration {
	my $conf_obj = shift;
	
	$conf_obj->{fields} = {
		status    => 'status',
		estimated => 'timeoriginalestimate',
		remaining => 'timeestimate',
		logged    => 'timespent'
	};
	
	return;
}

sub check_for_updates {
	my $conf_obj     = shift;
	my $current_file = convert_file_name_to_relative($0);
	my $new_file     = "tmp.txt";
	my $git          = WWW::Github::Files->new(
		author => 'anrodkin',
		resp   => 'jira_metrics',
		branch => 'master'
	);
	
	next if defined $conf_obj->{settings}->{check_script_for_updates} and not $conf_obj->{settings}->{check_script_for_updates};
	
	print "\nChecking for updates: ";
	
	my $file = $git->get_file("/$current_file");
	
	open my $output_file_handle, ">", "$new_file" or croak $!;
	print $output_file_handle $file;
	close $output_file_handle;
	
	my $checksum_for_current_file = crc32_file_hex($current_file);
	my $checksum_for_new_file     = crc32_file_hex($new_file);
	
	if("$checksum_for_current_file" ne "$checksum_for_new_file") {
		print "Update is available\n";
		
		copy("$new_file", "$current_file");
		remove("$new_file");
		
		print "\n\n\nScript was updated. Please rerun script\n";
		exit 0;
	}
	
	print "No updates\n";
	
	return;
}

sub check_and_install_required_modules {
	my $required_modules = shift;
	
	my @not_installed = grep {!CPAN->has_inst("$_")} keys %{$required_modules};

	foreach my $module (@not_installed) {
		print "Module missed: $module. Installing..\n";
		CPAN->install("$module");
	}
	
	if($#not_installed >= 0) {
		print "\n\n\nMissed modules were installed. Please rerun script.\n";
		exit 0;
	}
	
	return;
}

sub import_required_modules {
	my $modules = shift;
	
	foreach my $module (keys %{$modules}) {
		print "Import module $module: ";
		
		if($modules->{$module}) {
			my $import_without_params = grep { not defined $_ } @{$modules->{$module}};
			my @import_with_params    = grep {     defined $_ } @{$modules->{$module}};
			
			if($import_without_params) {
				$module->import();
			}
			
			if($#import_with_params >= 0) {
				$module->import(@import_with_params);
			}
		} else {
			$module->import();
		}
		
		print "OK\n";
	}
	
	return;
}

1;