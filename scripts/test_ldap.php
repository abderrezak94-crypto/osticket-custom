<?php
// CONFIGURATION LDAP (exactement comme dans ton plugin osTicket)
$ldap_host = "ldap://192.168.11.80"; // ✅ corrigé !
$ldap_user = "OSTICKET-METIDJ\\Administrator";
$ldap_pass = "Password01";
$ldap_dn   = "DC=osticket-metidji,DC=com"; // Search Base exact

// Connexion LDAP
$ds = ldap_connect($ldap_host);
ldap_set_option($ds, LDAP_OPT_PROTOCOL_VERSION, 3);
ldap_set_option($ds, LDAP_OPT_REFERRALS, 0);

if (@ldap_bind($ds, $ldap_user, $ldap_pass)) {
    echo "✅ Connexion LDAP réussie";
    ldap_unbind($ds);
} else {
    echo "❌ Échec de connexion LDAP";
}
?>