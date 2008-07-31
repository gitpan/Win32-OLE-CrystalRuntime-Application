# -*- perl -*-

use Test::More tests => 7;

is(".", ".", "We at least ran one test!");

SKIP: {
  eval {require Win32::OLE};
  my $gotWin32=not $@;
  skip qq{Package "Win32::OLE" is not available.}, 6 unless $gotWin32;
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
