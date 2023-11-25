<?php
defined('MOODLE_INTERNAL') || die();

function xmldb_qformat_xlsxtable_upgrade($oldversion) {

    if (get_config('converter_url', 'qformat_xlsxtable') !== false) {
        unset_config('converter_url', 'qformat_xlsxtable');
        unset_config('registration_url', 'qformat_xlsxtable');
        unset_config('username', 'qformat_xlsxtable');
        unset_config('password', 'qformat_xlsxtable');
    }

    return true;
}
