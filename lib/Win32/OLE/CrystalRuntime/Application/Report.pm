package Win32::OLE::CrystalRuntime::Application::Report;
use strict;
use Win32::OLE;
use Win32::OLE::Variant qw{VT_BOOL};
use constant True => Win32::OLE::Variant->new(VT_BOOL, 1);
use constant False=> Win32::OLE::Variant->new(VT_BOOL, 0);

BEGIN {
  use vars qw($VERSION);
  $VERSION     = '0.08';
}

=head1 NAME

Win32::OLE::CrystalRuntime::Application::Report - Perl Interface to the Crystal Report OLE Object

=head1 SYNOPSIS

  use Win32::OLE::CrystalRuntime::Application;
  my $application=Win32::OLE::CrystalRuntime::Application->new;
  my $report=$application->report(filename=>$filename);
  $report->export(format=>"pdf", filename=>"export.pdf");

=head1 DESCRIPTION

This package is a wrapper around the OLE object for a Crystal Report.

=head1 USAGE

=head1 CONSTRUCTOR

You must construct this object from a L<Win32::OLE::CrystalRuntime::Application> object as the ole object is constructed at the same time and is generated from the application->ole object.

  use Win32::OLE::CrystalRuntime::Application;
  my $application=Win32::OLE::CrystalRuntime::Application->new;
  my $report=$application->report(filename=>$filename);


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
  my $self = shift();
  %$self=@_;
  if (-r $self->filename) {
    $self->{'ole'}=$self->parent->ole->OpenReport($self->filename, 1);
    die(Win32::OLE->LastError) if Win32::OLE->LastError;
    die("Error: Cannot create OLE report object") unless ref($self->ole) eq "Win32::OLE";
  } else {
    die(sprintf(qq{Error: Cannot read file "%s".}, $self->filename));
  }
  $self->ole->DiscardSavedData();
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->{'DisplayProgressDialog'} = False;
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->{'MorePrintEngineErrorMessages'} = False;
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->{'EnableParameterPrompting'} = False;
  die(Win32::OLE->LastError) if Win32::OLE->LastError;
}

=head2 filename

Returns the name of the report filename. This value is read only after object construction.

  my $filename=$report->filename;

Set on construction

  my $report=Win32::OLE::CrystalRuntime::Application::Report->new(
               filename=>$filename,
             );

=cut

sub filename {
  my $self=shift;
  return $self->{'filename'};
}

=head2 ole

Returns the OLE report object.  This object is the Win32::OLE object that was constructed during initialization from the $application->report() method.

=cut

sub ole {
  my $self=shift;
  return $self->{'ole'};
}

=head2 parent

Returns the parent application object for the report.

  my $application=$report->parent;

Set on construction in the $application->report method.

  my $report=Win32::OLE::CrystalRuntime::Application::Report->new(
               parent=>$application
             );

=cut

sub parent {
  my $self=shift;
  return $self->{'parent'};
}

=head2 setParameters

Sets the report parameters.

  $report->setParameters($key1=>$value1, $key2=>$value2, ...);  #Always pass values as strings and convert in report
  $report->setParameters(%hash);

=cut

sub setParameters {
  my $self=shift;
  my $hash={@_};
  foreach my $index (1 .. $self->ole->ParameterFields->Count) {
    die(Win32::OLE->LastError) if Win32::OLE->LastError;
    my $key=$self->ole->ParameterFields->Item($index)->ParameterFieldName;
    die(Win32::OLE->LastError) if Win32::OLE->LastError;
    #printf qq{Setting Parameter: "%s" => "%s"\n}, $key, $hash->{$key};
    if (defined $hash->{$key}) {
      $self->ole->ParameterFields->Item($index)->AddCurrentValue($hash->{$key});
      die(Win32::OLE->LastError) if Win32::OLE->LastError;
    } else {
      warn(sprintf(qq{Warning: Report Parameter "%s" is not defined.}, $key));
    }
  }
}

=head2 export

Saves the report in the specified format to the specified filename.

  $report->export(filename=>"report.pdf");  #default format is pdf
  $report->export(format=>"pdf", filename=>"report.pdf");
  $report->export(formatid=>31, filename=>"report.pdf"); #pass format id directly

=cut

sub export {
  my $self=shift;
  my $opt={@_};
  my $formatid=$opt->{'formatid'} ||
                 $self->FormatID->{$opt->{'format'}||'pdf'};
  die("Error: export method requires a valid format.") unless $formatid;
  die("Error: export method requires a filename.") unless $opt->{'filename'};

  $self->ole->ExportOptions->{'DestinationType'} = 1;         #1=>filesystem
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->ExportOptions->{'FormatType'} = $formatid;
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->ExportOptions->{'DiskFileName'} = $opt->{'filename'};
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->ExportOptions->{'HTMLFileName'} = $opt->{'filename'};
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->ExportOptions->{'XMLFileName'} = $opt->{'filename'};
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->ExportOptions->{'PDFExportAllPages'} = True;
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  $self->ole->Export(False);
  die(Win32::OLE->LastError) if Win32::OLE->LastError;

  return $self;
}

=head2 FormatID

Returns a hash of common format extensions and CRExportFormatType IDs.  Other formats are supported with export(formatid=>$id)

  my $hash=$report->FormatID;           #{pdf=>31, xls=>36};
  my @orderedlist=$report->FormatID;    #(pdf=>31, xls=>36, ...)

=cut

sub FormatID {
  my $self=shift;
  my @data=qw{
              pdf 31
              xls 36
              doc 14
              csv 5
              rpt 1
              rtf 35
              htm 24
              html 32
              txt 10
              tsv 6
              xml 37
             };
  return wantarray ? @data : {@data};
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

=cut

1;
