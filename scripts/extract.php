<?php
$phar = new Phar('fr.phar');
$phar->extractTo('fr_lang', true); // Extrait dans le dossier fr_lang
echo "Extraction terminée !";
?>