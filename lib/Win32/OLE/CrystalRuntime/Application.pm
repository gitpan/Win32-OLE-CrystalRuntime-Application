package Win32::OLE::CrystalRuntime::Application;
use strict;
use Win32::OLE::CrystalRuntime::Application::Report;
use Win32::OLE;

BEGIN {
  use vars qw($VERSION);
  $VERSION='0.05';
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
  $report->export(type=>"pdf", filename=>"export.pdf");

=head1 DESCRIPTION

=head1 USAGE

=head1 CONSTRUCTOR

=head2 new

  my $application=Win32::OLE::CrystalRuntime::Application->new();

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

=head2 initialize

=cut

sub initialize {
  my $self = shift();
  %$self=@_;
  $self->ProgramID("CrystalRuntime.Application") unless defined $self->ProgramID;
}

=head2 ProgramID

Sets or returns the Program ID which defaults to "CrystalRuntime.Application".  You may want to specify the version if you have multiple objects in your environment.

  $application->ProgramID("CrystalRuntime.Application.11");  #Requires version 11

OR

  my $application=Win32::OLE::CrystalRuntime::Application->new(ProgramID=>"CrystalRuntime.Application.11");

=cut

sub ProgramID {
  my $self=shift;
  $self->{'ProgramID'}=shift if @_;
  return $self->{'ProgramID'};
}

=head2 ole

Set or Returns the OLE Application object.  This object is a Win32::OLE object that was created with a Program ID of "CrystalRuntime.Application.11"

=cut

sub ole {
  my $self=shift;
  $self->{'ole'}=shift if @_;
  unless (ref($self->{'ole'}) eq "Win32::OLE") {
    $self->{'ole'}=Win32::OLE->CreateObject($self->ProgramID);
    die(Win32::OLE->LastError) if Win32::OLE->LastError;
  }
  die("Error: Could not create the Win32::OLE object.") unless defined $self->{'ole'};
  return $self->{'ole'};
}

=head2 report

Returns a report object.

=cut

sub report {
  my $self=shift;
  my $report=Win32::OLE::CrystalRuntime::Application::Report->new(parent=>$self, @_);
  return $report;
}

=head1 BUGS

=head1 SUPPORT

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
