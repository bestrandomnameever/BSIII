#dsquery * "cn=schema,cn=configuration,dc=iii,dc=hogent,dc=be" -filter "(ldapdisplayname=cn)" -attr name

use Win32::OLE qw(in);
use bind_object;
use Data::Dumper;

@ARGV = ("cn");
@ARGV == 1 or die "geef de ldapdisplayName als enig argument !\n";

my $ldapdisplayname = $ARGV[0];


my $ADOconnection = Win32::OLE->new("ADODB.Connection");
   $ADOconnection->{Provider} = "ADsDSOObject";
   if ( uc($ENV{USERDOMAIN}) ne "III") { #niet ingelogd op het III domein
      $ADOconnection->{Properties}->{"User ID"}          = "Cedric Vanhaverbeke";
      $ADOconnection->{Properties}->{"Password"}         = "Cedric Vanhaverbeke";
      $ADOconnection->{Properties}->{"Encrypt Password"} = True;
   }
   $ADOconnection->Open();                                       # mag je niet vergeten
my $ADOcommand = Win32::OLE->new("ADODB.Command");
   $ADOcommand->{ActiveConnection}      = $ADOconnection;        # verwijst naar het voorgaand object
   $ADOcommand->{Properties}->{"Page Size"} = 20;

my $RootDSE=bind_object::bind_object("RootDSE"); #serverless Binding
   $RootDSE->getinfo();


my $schema  = bind_object::bind_object( $RootDSE->Get("schemaNamingContext") );

my $sBase  = $schema->{adspath};
my $sFilter = "(ldapdisplayname=$ldapdisplayname)";
my $sAttributes = "cn"; #of name
my $sScope      = "subtree";
   $ADOcommand->{CommandText} = "<$sBase>;$sFilter;$sAttributes;$sScope";
   $ADOcommand->{Properties}->{"Sort On"} = "cn";

my $ADOrecordset = $ADOcommand->Execute();


until ( $ADOrecordset->{EOF} )  {
   print $ADOrecordset->Fields("cn")->{Value},"\n";
   $ADOrecordset->MoveNext();
}
   $ADOrecordset->Close();
   $ADOconnection->Close();


#Je kan dit ook als volgt ophalen, zonder LDAP-query (zie reeks 6 oefening 39)
my $abstracteKlasse  = bind_object::bind_object( "schema/$ldapdisplayname" );
print "\nHet overeenkomstig reeel object heeft cn=",$abstracteKlasse->get("cn");