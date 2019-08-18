use bind_object;
use Data::Dumper;

my $User = bind_object::bind_object("CN=Users,DC=iii,DC=hogent,DC=be");
%constanten = bind_object::get_constanten();
foreach(in $User){
    printf("%s staat %s\n", "ADS_UF_ACCOUNTDISABLE", $_->{AccountDisabled} ? "aan" : "uit" );
    printf("in UserAccountControl staat het ADS_UF_ACCOUNTDISABLE bit  %s\n", $_->{UserAccountControl} & $constanten{ADS_UF_ACCOUNTDISABLE} ? "aan" : "uit");
    printf("%s staat %s\n", "ADS_UF_PASSWD_NOTREQD", $_->{PasswordRequired} ? "uit" : "aan");
    # Er staat not required in de bit, dus aan en uit omdraaien, kijk naar de oplossing voor het juiste
    printf("in UserAccountControl staat het ADS_UF_PASSWD_NOTREQD bit  %s\n", $_->{UserAccountControl} & !$constanten{ADS_UF_PASSWD_NOTREQD} ? "uit" : "aan");
    print "\n\n\n"
}