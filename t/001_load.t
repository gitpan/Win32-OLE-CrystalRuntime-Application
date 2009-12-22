# -*- perl -*-

use Test::More tests => 7;

is(".", ".", "We at least ran one test!");

my $skip=0;
eval {require Win32::OLE};
if ($@) { #no Win32::OLE
  $skip=1;
} else {  #have Win32::OLE
  my $obj=Win32::OLE->CreateObject(qq{CrystalRuntime.Application});
  $skip=Win32::OLE->LastError; #either "0" or "Invalid class string"
  warn($skip) if $skip;
  undef $obj;
}

SKIP: {
  skip qq{Win32::OLE and CrystalRuntime.Application are not available.}, 6
    if $skip;

  use_ok( 'Win32::OLE::CrystalRuntime::Application' );
  use_ok( 'Win32::OLE::CrystalRuntime::Application::Report' );
  my $application=Win32::OLE::CrystalRuntime::Application->new;
  isa_ok ($application, 'Win32::OLE::CrystalRuntime::Application');
  isa_ok ($application->ole, 'Win32::OLE');

  my $file="hello.rpt";
  my $filename=$file;
  foreach my $path (qw{.. t .} ) {
    $filename="$path/$file";
    last if -r $filename;
  }
  my $gotFile=-r $filename;
  skip qq{File "$file" is not readable.}, 2 unless $gotFile;
  my $report=$application->report(filename=>$filename);
  isa_ok ($report, 'Win32::OLE::CrystalRuntime::Application::Report');
  isa_ok ($report->ole, 'Win32::OLE');
}
