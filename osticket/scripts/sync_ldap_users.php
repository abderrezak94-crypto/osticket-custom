<?php

// Configuration
$ldap_host = "ldap://192.168.11.80";
$ldap_user = "OSTICKET-METIDJ\\Administrator";
$ldap_pass = "Password01";
$ldap_base_dn = "DC=osticket-metidji,DC=com";

$db_host = "localhost";
$db_user = "osticket";
$db_pass = "Password01";
$db_name = "osticket";

// Connexion LDAP
$ldap_conn = ldap_connect($ldap_host);
ldap_set_option($ldap_conn, LDAP_OPT_PROTOCOL_VERSION, 3);
ldap_set_option($ldap_conn, LDAP_OPT_REFERRALS, 0);

if (!@ldap_bind($ldap_conn, $ldap_user, $ldap_pass)) {
    die("❌ Échec de la connexion LDAP");
}
echo "✅ Connexion LDAP réussie\n";

// Connexion MySQL
$mysqli = new mysqli($db_host, $db_user, $db_pass, $db_name);
if ($mysqli->connect_error) {
    die("❌ Échec MySQL: " . $mysqli->connect_error);
}

// Requête LDAP
$filter = "(mail=*)";
$attributes = array("cn", "mail");
$search = ldap_search($ldap_conn, $ldap_base_dn, $filter, $attributes);
$entries = ldap_get_entries($ldap_conn, $search);

for ($i = 0; $i < $entries["count"]; $i++) {
    $entry = $entries[$i];

    $name = isset($entry["cn"][0]) ? $entry["cn"][0] : '';
    $email = isset($entry["mail"][0]) ? strtolower($entry["mail"][0]) : '';

    if (!$email || !$name) continue;

    // Vérifie si l'utilisateur existe
    $stmt = $mysqli->prepare("SELECT user_id FROM ost_user_email WHERE address = ?");
    $stmt->bind_param("s", $email);
    $stmt->execute();
    $stmt->bind_result($user_id);
    $exists = $stmt->fetch();
    $stmt->close();

    if ($exists) {
        // ✅ Met à jour la source
        $stmt = $mysqli->prepare("UPDATE ost_user__cdata SET source = 'Active Directory' WHERE user_id = ?");
        $stmt->bind_param("i", $user_id);
        $stmt->execute();
        $stmt->close();

        // ✅ Vérifie s'il a un compte
        $stmt = $mysqli->prepare("SELECT id FROM ost_user_account WHERE user_id = ?");
        $stmt->bind_param("i", $user_id);
        $stmt->execute();
        $stmt->store_result();

        if ($stmt->num_rows > 0) {
            // ✅ Met à jour le statut à Actif
            $stmt->close();
            $stmt = $mysqli->prepare("UPDATE ost_user_account SET status = 1 WHERE user_id = ?");
            $stmt->bind_param("i", $user_id);
            $stmt->execute();
            $stmt->close();
        } else {
            $stmt->close();
            // ✅ Crée le compte actif
            $backend = 'ldap.client.p2i1';
            $stmt = $mysqli->prepare("INSERT INTO ost_user_account (user_id, status, backend, registered) VALUES (?, 1, ?, NOW())");
            $stmt->bind_param("is", $user_id, $backend);
            $stmt->execute();
            $stmt->close();
        }

        echo "✅ Mise à jour : $name <$email>\n";
        continue;
    }

    // Nouveau compte utilisateur
    $stmt = $mysqli->prepare("INSERT INTO ost_user (default_email_id, name, created, updated) VALUES (0, ?, NOW(), NOW())");
    $stmt->bind_param("s", $name);
    $stmt->execute();
    $user_id = $stmt->insert_id;
    $stmt->close();

    // Email principal
    $stmt = $mysqli->prepare("INSERT INTO ost_user_email (user_id, address, flags) VALUES (?, ?, 1)");
    $stmt->bind_param("is", $user_id, $email);
    $stmt->execute();
    $email_id = $stmt->insert_id;
    $stmt->close();

    // Mettre à jour le default_email_id
    $stmt = $mysqli->prepare("UPDATE ost_user SET default_email_id = ? WHERE id = ?");
    $stmt->bind_param("ii", $email_id, $user_id);
    $stmt->execute();
    $stmt->close();

    // cdata avec source Active Directory
    $stmt = $mysqli->prepare("INSERT INTO ost_user__cdata (user_id, name, email, source) VALUES (?, ?, ?, 'Active Directory')");
    $stmt->bind_param("iss", $user_id, $name, $email);
    $stmt->execute();
    $stmt->close();

    // ✅ Création du compte osTicket avec statut Actif
    $backend = 'ldap.client.p2i1';
    $stmt = $mysqli->prepare("INSERT INTO ost_user_account (user_id, status, backend, registered) VALUES (?, 1, ?, NOW())");
    $stmt->bind_param("is", $user_id, $backend);
    $stmt->execute();
    $stmt->close();

    echo "✅ Utilisateur ajouté : $name <$email>\n";
}

ldap_unbind($ldap_conn);
$mysqli->close();


?>
