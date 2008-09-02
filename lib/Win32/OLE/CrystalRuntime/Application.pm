package Win32::OLE::CrystalRuntime::Application;
use strict;
use Win32::OLE::CrystalRuntime::Application::Report;
use Win32::OLE;

BEGIN {
  use vars qw($VERSION);
  $VERSION='0.08';
}

=head1 NAME

Win32::OLE::CrystalRuntime::Application - Perl Interface to the CrystalRuntime.Application OLE Object

=head1 SYNOPSIS

The ASP Version

    Dim oApp, oRpt
    Set oApp = Server.CreateObject("CrystalRuntime.Application")
    Set oRpt = oApp.OpenReport(vFilenameReport, 1)
    oRpt.DisplayProgressDialog = False
    oRpt.MorePrintEngineErrorMessages = False
    oRpt.EnableParameterPrompting = False
    oRpt.DiscardSavedData
    oRpt.ExportOptions.DiskFileName = vFilenameExport
    oRpt.ExportOptions.FormatType = 31                  '31=>PDF
    oRpt.ExportOptions.DestinationType = 1              '1=>filesystem
    oRpt.ExportOptions.PDFExportAllPages = True
    oRpt.Export(False)
    Set oRpt = Nothing
    Set oApp = Nothing

The perl Version

  use Win32::OLE::CrystalRuntime::Application;
  my $application=Win32::OLE::CrystalRuntime::Application->new;
  my $report=$application->report(filename=>$filename);
  $report->setParameters($key1=>$value1, $key2=>$value2);
  $report->export(format=>"pdf", filename=>"export.pdf");

=head1 DESCRIPTION

This package allows automation of generating Crystal Reports with Perl.  This package connects to the Crystal Runtime Application OLE object.  You MUST have a license for the Crystal Reports server-side component in order for this to work.

                                               Perl API       
                                                  |           
                                        +--------------------+
            Perl API                 +---------------------+ |
               |                  +----------------------+ | |
  +---------------------------+ +----------------------+ | | |
  |                           | |                      | | | |
  |  Perl Application Object  | |  Perl Report Object  | | | |
  |                           | |                      | | | |
  |       "ole" method        | |     "ole" method     | | | |
  |     +==============+      +-+   +==============+   | | | |
  |     |              |      | |   |              |   | | | |
  |     |  Win32::OLE  |      | |   |  Win32::OLE  |   | | | |
  |     |  Application |============|    Report    |   | | | |
  |     |    Object    |      | |   |    Object    |   | | |-+
  |     |              |      | |   |              |   | |-+
  |     +==============+      | |   +==============+   |-+
  +---------------------------+ +----------------------+ 

=head1 USAGE

  use Win32::OLE::CrystalRuntime::Application;
  my $application=Win32::OLE::CrystalRuntime::Application->new;
  my $report=$application->report(filename=>$filename);
  foreach my $index (1 .. $report->ole->Database->Tables->Count) {
    if ($report->ole->Database->Tables->Item($index)->DllName eq "crdb_oracle.dll") {
      $report->ole->Database->Tables->Item($index)->ConnectionProperties("Server")->{'Value'} = $database;
      $report->ole->Database->Tables->Item($index)->ConnectionProperties("User ID")->{'Value'} = $account;
      $report->ole->Database->Tables->Item($index)->ConnectionProperties("Password")->{'Value'} = $password;
    } 
  }
  $report->setParameters($key1=>$value1, $key2=>$value2);
  $report->export(format=>"pdf", filename=>"export.pdf");

=head1 CONSTRUCTOR

=head2 new

  my $application=Win32::OLE::CrystalRuntime::Application->new(
                    ProgramID=>"CrystalRuntime.Application", #default
                  );

=cut

sub new {
  my $this = shift();
  my $class = ref($this) || $this;
  my $self = {};
  bless $self, $class;
  $self->initialize(@_);
  return $self;
}

=head1 METHODS

=cut

sub initialize {
  my $self=shift;
  %$self=@_;
  $self->{'ProgramID'}="CrystalRuntime.Application"
    unless defined $self->ProgramID;
}

=head2 ProgramID

Returns the Program ID which defaults to "CrystalRuntime.Application".  You may want to specify the version if you have multiple objects in your environment.

  $application->ProgramID;

=cut

sub ProgramID {
  my $self=shift;
  return $self->{'ProgramID'};
}

=head2 ole

Set or Returns the OLE Application object.  This object is a Win32::OLE object that is created with a Program ID of "CrystalRuntime.Application"

=cut

sub ole {
  my $self=shift;
  $self->{'ole'}=shift if @_;
  unless (ref($self->{'ole'}) eq "Win32::OLE") {
    $self->{'ole'}=Win32::OLE->CreateObject($self->ProgramID);
    die(Win32::OLE->LastError) if Win32::OLE->LastError;
  }
  die("Error: Could not create the Win32::OLE object.")
    unless defined $self->{'ole'};
  return $self->{'ole'};
}

=head2 report

Constructs a report object which is a L<Win32::OLE::CrystalRuntime::Application::Report>.

  my $report=$application->report(filename=>$filename);

=cut

sub report {
  my $self=shift;
  my $report=Win32::OLE::CrystalRuntime::Application::Report->new(parent=>$self, @_);
  return $report;
}

=head1 BUGS

=head1 SUPPORT

Please try Business Objects.

=head1 AUTHOR

    Michael R. Davis
    CPAN ID: MRDVT
    STOP, LLC
    domain=>stopllc,tld=>com,account=>mdavis
    http://www.stopllc.com/

=head1 COPYRIGHT

This program is free software licensed under the...

	The BSD License

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

Crystal Reports XI Technical Reference Guide - http://support.businessobjects.com/documentation/product_guides/boexi/en/crxi_Techref_en.pdf

L<Win32::OLE>

=cut

1;
