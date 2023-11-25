<?php
defined('MOODLE_INTERNAL') || die();

$plugin->version   = 2023112500;
$plugin->maturity  = MATURITY_STABLE;
$plugin->component  = 'moodle2xlsx';
$plugin->release  = '0.0.1 (Build: 2023112500)';
$plugin->requires = 2018051700;  // Requires Moodle 3.5 or later.
$plugin->dependencies = array('booktool_xlsximport' => 2021083100);
