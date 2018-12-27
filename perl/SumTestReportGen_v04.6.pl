#!/usr/bin/env perl
##############################################################################
# File   : SumTestReportGen.pl
# Author : tuyendinh  <tuyendvu@geservs.com>
# Version: 4.6
##############################################################################
use warnings;
use strict;
use HTML::TableExtract;
use Data::Dumper;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Office';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
use File::Basename;
use File::Spec::Functions qw(rel2abs);
#use File::Spec;
use Getopt::Long qw(GetOptions);
use utf8;
#use warnings  qw(FATAL utf8);    # fatalize encoding faults
use Encode qw/encode decode/;
binmode(STDOUT, ":utf8");
use Getopt::Long;
use Pod::Usage;

#####################
use constant {
    #table zero
    MATCH_J      => '全体の一致',
    MATCH_E      => '全体の一致',
    RESULT_J     => '全体の合否',
    RESULT_E     => 'All Output Comparison',
    NO_TESTCASE_J	 => '総テストベクタ数',
    NO_TESTCASE_E  => 'Total of Test Vector',
    #table one
    # Colum name contain  CSV name 
    TABLE1_CSV_FILE_COLUMN_J   => 'CSVファイル',
    TABLE1_CSV_FILE_COLUMN_E   => 'CSV File name',
    # Column name contain function
    TABLE1_FUNCTION_TEST_COLUMN_J => '関数名',
    TABLE1_FUNCTION_TEST_COLUMN_E => 'Function',
    #table two
    TABLE2_FUNCTION_TEST_COLUMN_J => '関数',
    TABLE2_FUNCTION_TEST_COLUMN_E => 'Function',
    TABLE2_C0_COLUMN_J    => 'C0網羅率',
    TABLE2_C0_COLUMN_E    => 'C0 Coverage',
    TABLE2_C1_COLUMN_J    => 'C1網羅率',
    TABLE2_C1_COLUMN_E    => 'C1 Coverage',
    TABLE2_MCDC_COLUMN_J	=> 'MC/DC',
    TABLE2_MCDC_COLUMN_E	=> 'MC/DC',
    TABLE2_COL_LOG_J=> 'カバレッジログファイル',
    TABLE2_COL_LOG_E=> 'Coverage Log',
    #table three
    ROW_WINAMS  => 	0,
    ROW_CS2     =>	1,
    ROW_XAIL    =>	2,
    ROW_SIM     =>	3,
    ROW_OMF     =>	4,
    # number of row in a table
    #ENCODE_FORM => 'shiftjis'
    ENCODE_FORM => 'x-sjis'

};
our $MATCH = MATCH_J;
our $RESULT = RESULT_J;     
our $TABLE1_CSV_FILE_COLUMN = TABLE1_CSV_FILE_COLUMN_J;
our $NO_TESTCASE= NO_TESTCASE_J;
our $TABLE1_FUNCTION_TEST_COLUMN = TABLE1_FUNCTION_TEST_COLUMN_J;
our $TABLE2_FUNCTION_TEST_COLUMN = TABLE2_FUNCTION_TEST_COLUMN_J;
our $TABLE2_C0_COLUMN = TABLE2_C0_COLUMN_J;
our $TABLE2_C1_COLUMN = TABLE2_C1_COLUMN_J;
our $TABLE2_MCDC_COLUMN = TABLE2_MCDC_COLUMN_J;
our $TABLE2_COL_LOG = TABLE2_COL_LOG_J;

# Global variable
#our $SummaryReportPath    = "01_Summary_Report";
#our $TestSpecPath         = "02_Test_Spec";
#our $TestResultPath       = "03_Test_Result";
#our $IssueReportsPath     = "04_Issue_Reports";
#our $ClarificationsPath   = "05_Clarifications";
#our $SourceCodePath       = "06_Source_Code";

our $SummaryReportPath    = "00_Summary_Report";
our $TestSpecPath         = "01_Test_Spec";
our $TestResultPath       = "02_Test_Result";
our $IssueReportsPath     = "03_Issue_Reports";
our $ClarificationsPath   = "04_Clarifications";
our $SourceCodePath       = "05_Source_Code";
# Variable contain relative path to file 
our %CoverageLogPath;
our %hCsvResultPath;
our %hTestReportPath;use Getopt::Long qw(GetOptions);
our %hCsvFilePath;
our %hSpecFilePath;
our %hProblemPath;
our %hSrcFileReport;
our $Excel;
our $SummaryTestReport;
my $RootPath;
my $ExTemplate=join("\\",dirname(rel2abs($0)),"Template.xlsx");
my $TestLog;
my $help;
my $debug;
my $ErrorLog;
my $AddLinkTestSpec;
our $issueReportLink2TestSpec = 0;
# Folder contain: 00_Summary_Report, 01_Test_Spec, 02_Test_Result, 03_Issue_Reports, 04_Clarifications, 05_Source_Code
pod2usage("SumTestReport.exe: No input information") if ((@ARGV == 0) && (-t STDIN));
GetOptions('dir=s' => \$RootPath,    
    'xlsx=s'   => \$ExTemplate, # Excel Template
    'testlog|tl=s'=>\$TestLog,    # TestLog
    'ts'=>\$AddLinkTestSpec,    # TestLog
    'help|h|?'   => \$help,
    'debug|d'    => \$debug) or (pod2usage(2) && exit);
pod2usage(1) if $help;
unless (-d $RootPath)
{
    print ("Invalid path :$RootPath\n");
    exit;
}
unless (-f $ExTemplate)
{
    print ("Invalid Excel template file :$ExTemplate\n");
    exit;
}
my $now = time;
TestReportParser($RootPath, $TestLog);
my $time_parse_htm = time -$now;
printf("\nParsing html time: %02d:%02d:%02d\n\n", int($time_parse_htm/ 3600), int(($time_parse_htm% 3600) / 60), int($time_parse_htm % 60));
GetPathTestSpec($RootPath);
GetIssueReportPath($RootPath);
#debug_log(Dumper(%CoverageLogPath));
WriteData2Excel($RootPath);
#TODO
#CheckConsistency();
Write2File($ErrorLog);
$now = time - $now;
printf("\nExecution time: %02d:%02d:%02d\n\n", int($now / 3600), int(($now % 3600) / 60), int($now % 60));

######dd######## Sub definition #########################################

# TODO Create excel file in the case no templete.
sub CreateExcel
{
}

# Write data to excel
sub WriteData2Excel
{#{{{

    # Write data to excel file
    my $RootPath=shift;
    my $refBook=OpenExcel($ExTemplate);
    $$refBook->SaveAs(join('\\',$RootPath,$SummaryReportPath,'SummaryTestReport.xlsx'));
    my $Sheet=$$refBook->Worksheets('Function List') or die ("There are problem in opening worksheet\n");
    my $RowIndex=2;
    my $Index=1;
    my $IsReportIssue=0;
    # Reset text color to black
    $Sheet->Range("A:M")->EntireColumn->Font->{ColorIndex} = 1;
    $Sheet->Range("A:M")->Hyperlinks->Delete;
    $Sheet->Range("A:M")->Font->{Bold}       = "False";
    $Sheet->Range("A:M")->Font->{Italic}     = "False";
    $Sheet->Range("A:M")->Font->{Underline}  = -4142;
    $Sheet->Range("E:L")->EntireColumn->{HorizontalAlignment} = -4108;
    $Sheet->Range("E:L")->EntireColumn->{VerticalAlignment} = -41108; 
    $Sheet->Range("A1:M1")->Font->{Bold}       = "True";
    #Set column width and row height
    $Sheet -> Range("B:D")->EntireColumn->{ColumnWidth} = 60;
    $Sheet -> Range("B:D")->EntireColumn->{RowHeight}   = 30;
    foreach my $SrcFile (sort keys(%hSrcFileReport))
    {
        my @aData=@{$hSrcFileReport{$SrcFile}};
        foreach my $refRowData (@aData)
        {
            #unshift @$refRowData, $Index;
            # Format column
            $Sheet->Range("A$RowIndex:D$RowIndex")->{NumberFormat} = "\@";  # Text 
            $Sheet->Range("E$RowIndex:F$RowIndex")->{NumberFormat} = "#";   # Number
            $Sheet->Range("G$RowIndex:I$RowIndex")->{NumberFormat} = "0%";   # Percent 
            $Sheet->Range("J$RowIndex:M$RowIndex")->{NumberFormat} = "\@";   # Text
            # Write data from column A-K
            # ($TestID, $SrcFile, $FunctionName, $NoTestCase,$NoTestPass, $C0, $C1, $MCDC, $Result, $Match);

            $Sheet->Range("A$RowIndex:A$RowIndex")->{Value} = $Index;
            $Sheet->Range("B$RowIndex:K$RowIndex")->{Value} = $refRowData;
            # Create hyperlink
            # The first elment of @$refRowData is the TestID (the CSV file name without extention)
            my $TestID=$refRowData->[0];
            unless (defined($hCsvFilePath{$TestID}))
            {
                print ("ERROR: The $TestID.csv does not exist\n");
                $ErrorLog .="ERROR: The $TestID.csv does not exist\n";
            }
            else 
            {
                # Csv file
                $Sheet->Hyperlinks->Add($Sheet->Range("E$RowIndex"),join('\\','..',$hCsvFilePath{$TestID}))
            }

            unless (defined($hSpecFilePath{$TestID}))
            {
                print ("ERROR: The $TestID.xlsx does not exist\n");
                $ErrorLog .="ERROR: The $TestID.xlsx does not exist\n";
            }
            else
            {
                # Add link to test specification
                if ($AddLinkTestSpec)
                {
                    $Sheet->Range("N$RowIndex")->{Value} = 'link';
                    my $tsLink;
                    $tsLink = join('\\','..',$hSpecFilePath{$TestID})."#Cover!A1";
                    $Sheet->Hyperlinks->Add($Sheet->Range("N$RowIndex"),$tsLink);
                }
            }
            # Coverage
            $Sheet->Hyperlinks->Add($Sheet->Range("I$RowIndex"),join('\\','..',$CoverageLogPath{$TestID}));
            if ($refRowData->[7] eq 'N/A' or $refRowData->[7] eq '')
            {    # Link to C1
                $Sheet->Hyperlinks->Add($Sheet->Range("H$RowIndex"),join('\\','..',$CoverageLogPath{$TestID}));
            }
            else
            {    # Link to MCDC
                $Sheet->Hyperlinks->Add($Sheet->Range("I$RowIndex"),join('\\','..',$CoverageLogPath{$TestID}));
            }
            #Test Report
            $Sheet->Hyperlinks->Add($Sheet->Range("J$RowIndex"),join('\\','..',$hTestReportPath{$TestID}));

            delete $hCsvFilePath{$TestID};
            # Check issue report 
            unless ($refRowData->[5] eq '100%')
            {
                $Sheet->Range("G$RowIndex:G$RowIndex")->Font->{ColorIndex} = 3;
                $IsReportIssue =1;
            }
            unless ($refRowData->[6] eq '100%')
            {
                $Sheet->Range("H$RowIndex:H$RowIndex")->Font->{ColorIndex} = 3;
                $IsReportIssue =1; 
            }
            unless ($refRowData->[7] eq '100%' or ($refRowData->[9] eq 'N/A') and ($refRowData->[7] eq 'N/A'))
            {
                $Sheet->Range("I$RowIndex:I$RowIndex")->Font->{ColorIndex} = 3;
                $IsReportIssue =1; 
            }
            unless ($refRowData->[8] eq 'OK')
            {
                $Sheet->Range("J$RowIndex:J$RowIndex")->Font->{ColorIndex} = 3;
                $IsReportIssue =1;
            }
            unless (encode('utf8',$refRowData->[9]) eq encode('utf8','一致') or encode('utf8',$refRowData->[9]) eq encode('utf8','Match') or (($refRowData->[9] eq 'N/A') and ($refRowData->[7] eq 'N/A')))
            {
                $Sheet->Range("K$RowIndex:K$RowIndex")->Font->{ColorIndex} = 3;
                $IsReportIssue =1;
            }
            # Check test case does not have expected result
            if (($refRowData->[4] ne $refRowData->[5]) and ($refRowData->[9] eq 'OK'))
            {
                my $temp = $refRowData->[4] - $refRowData->[5];
                warn ("ERROR: There are $temp test cases in the file $TestID.csv do not have expected result\n");
                $ErrorLog .= "ERROR: There are $temp test cases in the file $TestID.csv do not have expected result\n";
            }
            # checking issue report exist or not
            if($IsReportIssue eq 1)
            {
                if ($issueReportLink2TestSpec eq 0)
                {
                    unless(defined($hProblemPath{$TestID}))
                    {
                        $Sheet->Range("L$RowIndex:L$RowIndex")->{Value} = "Yes";
                        warn ("ERROR: No issue report for the test $TestID ?\n");
                        $ErrorLog .= "ERROR: No issue report for the test $TestID ?\n";
                    }
                }
                else
                {
                    $Sheet->Range("L$RowIndex:L$RowIndex")->{Value} = "Yes";
                    my $issueReportPath= join('\\','..',$hSpecFilePath{$TestID})."#IncidentReport!A1";
                    $Sheet->Hyperlinks->Add($Sheet->Range("L$RowIndex"),$issueReportPath);

                }

                $IsReportIssue =0;
            }
            else
            {
                $Sheet->Range("L$RowIndex:L$RowIndex")->{Value} = "No";
            }
            # There can be more issue rather than code coverage, test fail
            # # Write data to colum M (Issue report) even there is problem in C), C1, MCDC, MATCH, OK
            if(defined($hProblemPath{$TestID}))
            {
                $Sheet->Range("L$RowIndex:L$RowIndex")->{Value} = "Yes";
                $Sheet->Hyperlinks()->Add($Sheet->Range("L$RowIndex:L$RowIndex"),join('\\','..',$hProblemPath{$TestID}));
            }
            # Increase Index
            $Index++;
            $RowIndex++;
        }    
    }
    $Sheet -> Range("B:D")->EntireColumn->{Columns}->Autofit;
    $Sheet -> Range("B:D")->EntireColumn->{Rows}->Autofit;
    $$refBook->Save();
    $$refBook->Close();
    print "Finish Creating Test Summary Report: ".join('\\',$RootPath,$SummaryReportPath,'SummaryTestReport.xlsx')."\n";
    #$Excel->Quit();
    undef $$refBook;
    undef $Excel;
}#}}}
# Make sure all test has been executed
sub CheckConsistency
{#{{{
    my @csvFiles =(keys(%hCsvFilePath));
    if (keys(%hCsvFilePath) gt 0)
    {
        foreach (keys(%hCsvFilePath))
        {
            warn("ERROR: The Test report for the function $_ does not existed\n");
            $ErrorLog .= "ERROR: The Test report for the function $_ does not existed\n";
        }
    }
}#}}}
# Store data to hash hSrcFileReport
# Input: reference of array of TestReport.htm
#      reference of array of TestResultInfo.data 
# Ouput: reference of hash that contain data extract from TestReport.htm and TestResultInfor.data
sub TestReportParser
{
    my $RootPath=shift;
    my $Path=shift;
    unless (defined($Path))
    {
        $Path=join('\\', $RootPath,$TestResultPath);
    }
    my ($refTestCaseReportLog, $refTestResultInfo)=FindTestReport($Path);
    print("ERROR: The total file TestReport.htm and TestResutlInfo are not equal\n") unless(@$refTestCaseReportLog eq @$refTestResultInfo);
    for (my $i = 0; $i< @$refTestCaseReportLog;$i ++)
    {
        # Process htm file
        # Change language partern
        if ($refTestCaseReportLog->[$i] =~ /TestReport\.htm$/)
        {
            update_pattern(1);
        }
        else 
        {
            update_pattern(0);
        }

        my $ignoreTest=0;
        my $te = HTML::TableExtract->new();
        $te->parse_file($refTestCaseReportLog->[$i])||die("Can not open the file $refTestCaseReportLog->[$i]");
        my ($TestID, $SrcFile, $FunctionName, $Result, $Match, $NoTestCase, $CsvFile, $C0, $C1, $MCDC, $CoverageLog,$Startup);
        foreach my $ts ($te->tables) 
        {#{{{
            my @coords =  $ts->coords;
            # Table zero
            if ($coords[1] eq 0)
            {#{{{
                ($Result,$Match,$NoTestCase)=table_zero($ts);
            }#}}}
            # Table one
            elsif ($coords[1] eq 1)
            {#{{{
                ($CsvFile,$FunctionName)=table_one($ts);
            }#}}}
            # Table two
            elsif ($coords[1] eq 2)
            {#{{{
                ($C0,$C1,$MCDC,$CoverageLog, $SrcFile)=table_two($ts, $FunctionName);
            }#}}}
            # Table three
            elsif ($coords[1] eq 3)
            {#{{{
                table_three($ts);
            }#}}}
            else
            {#{{{
                print("ERROR: There are more than 4 tables in the test reports. The template format of the test report is not correct\n");
                $ignoreTest = 1;
                last;
            }#}}}
        }#}}}
        if($ignoreTest eq 0)
        {
            ($TestID = $CsvFile) =~ s/\.csv//g;
            if ($TestID eq "")
            {
                print ("ERROR: There is error in test report: $refTestCaseReportLog->[$i]\n");
                $TestID = "UNKNOWN";
            }
            # Process TestResultInfo file
            open my $file, '<', $refTestResultInfo->[$i]; 
            my $firstLine = <$file>; 
            close $file;
            my @tempData=split(',',$firstLine);
            my $NoTestPass=$tempData[-5];
            # Store link to TestReport.htm and CoverageLog
            my $RelTestReportPath=File::Spec->abs2rel($refTestCaseReportLog->[$i], $RootPath);
            my $CurDir=dirname($RelTestReportPath);
            $CoverageLogPath{$TestID}=join('\\',$CurDir,$CoverageLog);
            $hTestReportPath{$TestID}=$RelTestReportPath;
            # The result of each test case will be store to an array in order
            # NoTestCase, C0, C1, MDCD, Result, Match, CoverageLog, CsvFile
            my @aTestCaseData=($TestID, $SrcFile, $FunctionName, $NoTestCase,$NoTestPass, $C0, $C1, $MCDC, $Result, $Match);
            #my %hFuncReport= ("$FunctionName" =>\@aTestCaseData);
            push @{$hSrcFileReport{$SrcFile}},\@aTestCaseData;
        }
        else 
        {
            $ignoreTest = 0;
        }
    }
}

# Create a sub function to store path of test specification and test case, issue reports
# The key of hash is the file name (design ID)
sub GetPathTestSpec
{#{{{
    my $RootPath=shift;
    my $Path=join('\\', $RootPath,$TestSpecPath);
    my $refTestSpecList=FindFileType(join('\\',$Path,'01_Test_Design'),'.(xlsx|xlsm)$');
    my $refTestCaseList=FindFileType(join('\\',$Path,'02_Test_Case'),'.csv$');
    #my @aTestCase = @$refTestCaseList;
    #my @aTestSpec = @$refTestSpecList;
    #@aTestCase = map {$_ =~ s/([^\\\/]+)\.csv/$1/g; $_;} @aTestCase;
    #@aTestSpec = map {$_ =~ s/([^\\\/]+)\.xlsx/$1/g; $_; } @aTestSpec;
    #my @diffTestCase = diff_array( \@aTestCase, \@aTestSpec );
    #my @diffTestSpec = diff_array( \@aTestSpec, \@aTestCase );
    #if(scalar(@diffTestCase) gt 0)
    #{
    #    print "ERROR: Test case without test design\n";
    #    $ErrorLog .=  "ERROR: Test case without test design\n";
    #    foreach (@diffTestCase)
    #    {
    #            $ErrorLog .="$_\n";
    #            print "$_\n";
    #    }
    #}
    #if(scalar(@diffTestSpec) gt 0)
    #{
    #    print "ERROR: Test design without csv\n";
    #    $ErrorLog .=  "ERROR: Test design without csv\n";
    #    foreach (@diffTestSpec)
    #    {
    #            $ErrorLog .="$_\n";
    #            print "$_\n";
    #    }
    #    
    #}
    unless (scalar(@$refTestSpecList) eq (scalar(@$refTestCaseList)))
    {
        warn "ERROR: The total number of xlsx file and csv file are not equal TS: ".scalar(@$refTestSpecList). "\tTC: ".scalar(@$refTestCaseList)."\n";
        $ErrorLog .= "ERROR: The total number of xlsx file and csv file are not equal TS: ".scalar(@$refTestSpecList). "\tTC: ".scalar(@$refTestCaseList)."\n";
    }
    foreach (@$refTestCaseList)
    {
        if($_ =~ /([^\\\/]+)\.csv/)
        {
            $hCsvFilePath{$1}=File::Spec->abs2rel($_, $RootPath);
        }
        else
        {
            warn("ERROR: The naming convension of $_ is not correct\n");
            $ErrorLog .= "ERROR: The naming convension of $_ is not correct\n";
        }
    }
    foreach (@$refTestSpecList)
    {
        if($_ =~ /([^\\\/]+)\.(xlsm|xlsx)/)
        {
            my $StrTemp=$1;
            #$StrTemp =~ s/߿߿߿߿߿_//g;
            $hSpecFilePath{$StrTemp}=File::Spec->abs2rel($_, $RootPath);
        }
        else
        {
            warn("ERROR: The naming conversion of $_ is not correct");
            $ErrorLog .= "ERROR: The naming convension of $_ is not correct";
        }
    }
}
# Sub function store issue report path
# The key of hash is the file name (design ID)
sub GetIssueReportPath
{
    my $RootPath=shift;
    my $issueReportPath = join('\\',$RootPath,$IssueReportsPath);
    if (-d $issueReportPath)
    {
        my $refIssueList=FindFileType($issueReportPath,'.xlsx$');
        foreach (@$refIssueList)
        {
            if($_ =~ /([^\\\/]+)_Issue\.xlsx/)
            {    
                my $temp = $1;
                $temp =~ s/_Issue//g;
                $hProblemPath{$temp}=File::Spec->abs2rel($_, $RootPath);
            }
            else
            {
                warn("ERROR: The format name for issue report of $_ is not correct");
                $ErrorLog .= "ERROR: The format name for issue report of $_ is not correct ";
            }
        }
    }
    # Incident report is link to the sheet "incident report" in the test design"
    else
    {
        $issueReportLink2TestSpec = 1;
    }

}

sub table_zero
{#{{{
    # Need code check the table is valid or not. 
    my $ts=shift;
    my %hData;
    foreach my $row ($ts->rows)
    {
        (my $temp0 =($row->[0])) =~ s/^\s+|\s+$//g;
        #print "$temp0\n";
        #$temp0= encode('shiftjis',$temp0)."\n";
        #print decode('shiftjis',$temp0)."\n";
        #print "$temp0\n";
        (my $temp1 =$row->[1]) =~ s/^\s+|\s+$//g;
        $hData{$temp0}=decode(ENCODE_FORM,$temp1);
        #$hData{$temp0}=$temp1;
        #$temp0= encode('utf8',$temp0);
    }
    $hData{encode(ENCODE_FORM,$MATCH)}='N/A' unless (defined($hData{encode(ENCODE_FORM,$MATCH)}));
    return ($hData{encode(ENCODE_FORM,$RESULT)},$hData{encode(ENCODE_FORM,$MATCH)},$hData{encode(ENCODE_FORM,$NO_TESTCASE)});
}#}}}
sub table_one
{#{{{
    my $ts=shift;
    my $CsvFile;
    my $FunctionName;
    my $noData = 0;
    foreach ($ts->columns)
    {
        #Get CSV name
        if(trim($_->[0]) eq encode(ENCODE_FORM,$TABLE1_CSV_FILE_COLUMN))
        {    
            $noData ++;
            # Get only one csv file. Index start from 0 is header.
            $CsvFile=trim($_->[1]);
            $CsvFile =~ /([^\\\/]+\.csv)/g;
            $CsvFile = $1;
            next;
        }
        if (trim($_->[0]) eq encode (ENCODE_FORM,$TABLE1_FUNCTION_TEST_COLUMN))
        {
            $noData ++;
            # Get only one csv file. Index start from 0 is header.
            my @str=split('/',$_->[1]);
            # The function name can have form FileName.c/FunctionName 
            $FunctionName=trim($str[-1]);
            next;
        }
        # Stop process when get enough data
        last if ($noData eq 2);
    }
    return ($CsvFile,$FunctionName);

}#}}}
sub table_two
{#{{{
    my $ts=shift;
    my $FunctionName=shift;
    # There can be many functions in this table
    my $targetTestFunctionRow = 1;
    if ($FunctionName ne '')
    {
        my $index = 1;
        foreach my $col ($ts->columns)
        {
            if(trim($col->[0]) eq encode(ENCODE_FORM,$TABLE2_FUNCTION_TEST_COLUMN))
            {
                while($index < scalar(@{$col}))
                {
                    last if($col->[$index] =~ /$FunctionName/gi);
                    $index ++;
                }
                last;
            }
        }
        $targetTestFunctionRow = $index;
    }
    my ($C0, $C1, $MCDC, $CoverageLog, $SourceFile)=("UNDEFINED","UNDEFINED","UNDEFINED","UNDEFINED","UNDEFINED");
    foreach my $col ($ts->columns)
    {
        if (trim($col->[0]) eq encode (ENCODE_FORM,$TABLE2_C0_COLUMN))
        {
            $C0=trim($col->[$targetTestFunctionRow]);
            next;
        }
        if (trim($col->[0]) eq encode (ENCODE_FORM,$TABLE2_C1_COLUMN))
        {
            $C1=trim($col->[$targetTestFunctionRow]);
            next;
        }
        if (trim($col->[0]) eq encode (ENCODE_FORM,$TABLE2_MCDC_COLUMN))
        {
            $MCDC=trim($col->[$targetTestFunctionRow]);
            next;
        }
        if (trim($col->[0]) eq encode (ENCODE_FORM,$TABLE2_COL_LOG))
        {
            $CoverageLog=trim($col->[$targetTestFunctionRow]);
            my @str=split(/\\/,$CoverageLog);
            foreach (@str)
            {
                if ($_ =~ /\.c$/g)
                {
                    $SourceFile=$_;
                    last;
                }
            }
            next;
        }
    }
    $C0 = 'N/A' if ($C0 eq "UNDEFINED");
    $C1 = 'N/A' if ($C1 eq "UNDEFINED");
    $MCDC = 'N/A' if ($MCDC eq "UNDEFINED");
    return ($C0,$C1,$MCDC,$CoverageLog, $SourceFile);
}#}}}
sub table_three
{#{{{
    my $ts=shift;
    # hardcode
    my $winAMS= trim($ts->row(ROW_WINAMS)->[1]);
    my $CasePlayer=trim($ts->row(ROW_CS2)->[1]);
    my $XAIL=trim($ts->row(ROW_XAIL)->[1]);
    my $Simulator=trim($ts->row(ROW_SIM)->[1]);
    my $OMF=trim($ts->row(ROW_OMF)->[1]);
    return ($winAMS, $CasePlayer, $XAIL, $Simulator, $OMF);

}#}}}

#Remove white space
sub  trim 
{ #{{{
    my $s = shift;
    $s =~ s/^\s+|\s+$//g; 
    return $s; 
}#}}}

sub FindTestReport
{#{{{
    my @Folders=(shift);
    my @aTestCaseReport=();
    my @TestResultInfo=();
    foreach my $folder1 (@Folders)
    {
        opendir(DIR, $folder1) or die "Can not open $folder1 $!";
        my @FileList = readdir(DIR);
        closedir(DIR);
        foreach my $files1 (@FileList)
        {
            next if($files1 =~ /^\.\.?$/);
            my $FullPath=join('\\', $folder1, $files1);
            if ($files1 =~ /.htm$/ig)
            {
                push (@aTestCaseReport, $FullPath);
                push (@TestResultInfo, join('\\', $folder1, 'TestResultInfo.dat'));
                last;
            }
            if(-d $FullPath)
            {

                push (@Folders, $FullPath);
                #unshift (@Folders, $FullPath);
                next;
            }
        }
    }
    print "There are ". scalar(@aTestCaseReport) ." htm file\n";
    return \@aTestCaseReport, \@TestResultInfo;
}#}}}

# Find all file matching with pattern
# Input: Folder_Path, Pattern,
# Output: reference to array of files
sub FindFileType
{#{{{
    my ($Root, $Pattern)=@_;
    my @Folders=($Root);
    my @FileList;
    foreach my $folder1 (@Folders)
    {
        opendir(DIR, $folder1) or die "Can not open $folder1 $!";
        my @Files = readdir(DIR);
        closedir(DIR);
        foreach my $file (@Files)
        {
            next if($file =~ /^\.\.?$/);
            my $FullFilePath=join('\\', $folder1, $file);
            if ((-f $FullFilePath) && ($file =~ /$Pattern/))
            {
                push (@FileList, $FullFilePath);
                #next;
            }
            if(-d $FullFilePath)
            {
                push (@Folders, $FullFilePath);
                next;
            }
        }
    }
    return \@FileList;
}#}}}

sub OpenExcel
{#{{{
    my $ExTemplate=shift;
    # use existing instance if Excel is already running
    eval {$Excel = Win32::OLE->GetActiveObject('Excel.Application')};
    die "Excel not installed" if $@;
    unless (defined $Excel) {
        $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
            or die "Oops, cannot start Excel";
    }
    $Excel->{DisplayAlerts}=0;
    $Excel->{Visible} = 0;
    my $Book = $Excel->Workbooks->Open($ExTemplate) or die ("Can not open $ExTemplate\n");
    return (\$Book);
}#}}}

sub update_pattern
{
    # 1 => english, 0 -> Japanese
    my $lang = shift;
    if ($lang eq 1)
    {	
        $MATCH = MATCH_E;
        $RESULT = RESULT_E;     
        $TABLE1_CSV_FILE_COLUMN = TABLE1_CSV_FILE_COLUMN_E;
        $NO_TESTCASE= NO_TESTCASE_E;
        $TABLE1_FUNCTION_TEST_COLUMN = TABLE1_FUNCTION_TEST_COLUMN_E;
        $TABLE2_FUNCTION_TEST_COLUMN = TABLE2_FUNCTION_TEST_COLUMN_E;
        $TABLE2_C0_COLUMN = TABLE2_C0_COLUMN_E;
        $TABLE2_C1_COLUMN = TABLE2_C1_COLUMN_E;
        $TABLE2_MCDC_COLUMN = TABLE2_MCDC_COLUMN_E;
        $TABLE2_COL_LOG = TABLE2_COL_LOG_E;
    }
    else
    {
        $MATCH = MATCH_J;
        $RESULT = RESULT_J;     
        $TABLE1_CSV_FILE_COLUMN = TABLE1_CSV_FILE_COLUMN_J;
        $NO_TESTCASE= NO_TESTCASE_J;
        $TABLE1_FUNCTION_TEST_COLUMN = TABLE1_FUNCTION_TEST_COLUMN_J;
        $TABLE2_FUNCTION_TEST_COLUMN = TABLE2_FUNCTION_TEST_COLUMN_J;
        $TABLE2_C0_COLUMN = TABLE2_C0_COLUMN_J;
        $TABLE2_C1_COLUMN = TABLE2_C1_COLUMN_J;
        $TABLE2_MCDC_COLUMN = TABLE2_MCDC_COLUMN_J;
        $TABLE2_COL_LOG = TABLE2_COL_LOG_J;
    }
}

sub diff_array()
{#{{{
    my ($RefArray1,$RefArray2) = (@_);
    my %array2 = map {$_=>1} (@$RefArray2);
    my @onlyArray1 = grep { !$array2{$_} }(@$RefArray1); 
    return @onlyArray1;
}#}}}
sub Write2File
{#{{{
    my @string=@_;
    my $FileName= 'ErrorLog.txt';
    open (my $fh, '>', $FileName) or die ("Could not open file $FileName $!");
    foreach (@string)
    {
        if(defined ($_))
        {
            print $fh $_;
        }
    }
    close $fh;
}
sub debug_log
{
    if (defined($debug))
    {
        foreach (@_)
        {
            print "$_\n";
        }
    }
}

__END__#}}}

=head1 NAME
#{{{
    Using SumTestReportGen.exe to generate summary test report
#}}}
=head1 SYNOPSIS
#{{{
SumTestReportGen.exe [options] 

     Options:

       -dir (mandatory)  Path to root folder 
       -xlsx (optional) Path to excel templete file. 
       -help Message how to use tool
     Example: SumTestReportGen.exe -dir=C:\ISF_ME_2x_D0-2Pre5 -xlsx=C:\SumTestReportGen\Template.xlsx
     or
     Example: SumTestReportGen.exe -dir=C:\ISF_ME_2x_D0-2Pre5 #}}}
=cut  
