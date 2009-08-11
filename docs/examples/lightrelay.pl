#!/usr/bin/perl -w

use strict;
use warnings;
use POE qw(Wheel::FollowTail);
use DBI;
use DBD::mysql;

#### Config Options ####
my $port = "/dev/ttyS0";
my $bank = 0; # Bank number which relays on device belong to
our $color = 'green'; # Set initial color to Green
my $logfile = "/var/log/lightrelay.log"; # Where to output color changes to
my $filename = "/var/www/zm/events/2/dbgpipe.log"; # The file for this script to monitor
my $baud = 38400;
my $command = shift;
my $on_green = 108;
my $off_green = 100;
my $on_amber = 109;
my $off_amber = 101;
my $on_red = 110;
my $off_red = 102;
my $pid = "/tmp/lightrelay.pid";
# Database Options #
my $host = 'localhost';
my $database = 'lightrelay';
my $table = 'history';
my $user = 'lightrelay';
my $password = 'robot';
my $dsn = "dbi:mysql:$database:$host";


##!! Do not change anything below this line !!##
if (!$command || $command !~ /^(?:start|stop|restart|status|help)$/) {
 print("Usage: lightrelay.pl <start|stop|restart|status|help>\n");
 exit; 
}

if ($command eq "help") {
 print "start:\t\t Start the program.\n";
 print "stop:\t\t Stop the program.\n";
 print "restart:\t Restart the program.\n";
 print "status:\t\t Displays the current color.\n";
 print "help:\t\t This list.\n";
 exit;
}

if ($command eq "status") {
 system("/usr/bin/tail -1 $logfile");
 exit;
}

if ($command eq 'stop') {
 system("/bin/kill `cat /tmp/lightrelay.pid`");
 &turn_relays_off();
 exit;
}

if ($command eq 'restart') {
 system("/bin/kill `cat /tmp/lightrelay.pid`");
 system("/usr/local/bin/lightrelay.pl start &");
 exit;
}

if ($command eq 'start') {
 system("/bin/echo $$ > $pid") == 0 or warn "can not create pid file $pid";
 system("/bin/stty $baud < $port") == 0 or die "can not set device settings on $port";
 open(SERIALPORT, "+<", "$port") or die "can not open device $port";

POE::Session->create(
 inline_states => {
  _start => sub {
   $_[HEAP]{tailor} = POE::Wheel::FollowTail->new(
    Filename => "$filename",
    InputEvent => "got_log_line",
   );
  },
  got_log_line => sub {
   if ($_[ARG0] =~ /Green.*alarmed/ && $color eq 'red') # Color is red; green alarms...
    {
     &turned_color($off_red,$on_green,'green','Green');
    }
   elsif ($_[ARG0] =~ /LG.*alarmed/ && $color eq 'red') # Color is red; left green alarms...
    {
     &turned_color($off_red,$on_green,'green','Left Green');
    }
   elsif ($_[ARG0] =~ /Amber.*alarmed/ && $color eq 'green') # Color is green; amber alarms...
    {
     &turned_color($off_green,$on_amber,'amber','Amber');
    }
   elsif ($_[ARG0] =~ /Red.*alarmed/ && $color eq 'amber') # Color is amber; red alarms...
    {
     &turned_color($off_amber,$on_red,'red','Red');
    }
  },
 }
);
 POE::Kernel->run();
}


sub turned_color {
 $color = "$_[2]"; # Set color to Green, Amber or Red
 &send_signals($_[0],$_[1]); # Call sub to change relay status
 &log($_[3]);
}

sub send_signals {
 select((select(SERIALPORT), $|=1)[0]);
 print SERIALPORT chr(254); # Enter Command Mode
 print SERIALPORT chr($_[0]); # Deactivate Previous Relay
 print SERIALPORT chr($bank); # In Bank 1
 select(undef,undef,undef,.1); # Can not sleep() < 1, so we use select()'s timeout
 print SERIALPORT chr(254); # Enter Command Mode
 print SERIALPORT chr($_[1]); # Activate Current Relay
 print SERIALPORT chr($bank); # In Bank 1
}

sub log {
 my $state = $_[0];
 my $connect = DBI->connect($dsn,$user,$password) or warn "Unable to connect to mysql server $DBI::errstr\n";
 my $time = time();
 my $query = $connect->prepare("INSERT INTO history (color, epoch) VALUES ('$state', '$time()')");
 $query->execute();
}

sub turn_relays_off {
 open(SERIALPORT, "+<", "$port") or die "can not open device $port";
 print SERIALPORT chr(254);
 print SERIALPORT chr(29);
 close(SERIALPORT);
}
