<?php
/*
Plugin Name: Elementor Forms Statistics
Requires Plugins: elementor/elementor.php, elementor-pro/elementor-pro.php
Plugin URI: https://www.medienproduktion.biz/elementor-forms-extras/
Description: This plugin allows editors to view submissions received through Elementor forms. Additionally, a separate menu provides statistical analyses of the submissions, displayed as charts and tables.
Version: 1.1.2
Requires at least: 5.0
Tested up to: 6.9
Requires PHP: 7.2
Author: Nikolaos Karapanagiotidis
Author URI: https://www.medienproduktion.biz
Text Domain: elementor-forms-statistics

=== Changelog ===

= Version 1.1.2 – 13. Januar 2026 =
* Aktueller (laufender) Monat im Diagramm gestrichelt dargestellt.
* Export ist jetzt nur noch CSV (Excel-Auswahl entfernt).
* Abhängigkeit von Elementor Pro ergänzt.

= Version 1.1.1 – 08. Januar 2026 =
* Benutzerrollen für alle Menüeinträge (Statistik, Einstellungen, E-Mail Versand, Archiv, Export, Anfragen) einstellbar.
* Anfragen als Unterpunkt im Statistik-Menü, gesteuert über Rollenregeln.
* HTML-Export im E-Mail Versand unabhängig vom Export-Menü.

= Version 1.0.1 – 20. Dezember 2025 =
* Die automatische Bereinigung lässt sich jetzt per Dropdown (Deaktiviert | 1 Stunde | 1 Tag | 1 Monat | 1 Jahr) steuern, die Daten werden weiterhin erst nach Archivierung gelöscht.

= Version 1.1.0 – 02. Januar 2026 =
* Neuer Export-Bereich mit Tabs, Rollensteuerung und konfigurierbaren Export-Spalten inkl. Drag & Drop, Spaltennamen und Reihenfolge.
* Vorschau mit optionalen Spalten, Excel/CSV-Export, Merken des letzten Formulars und Formats sowie sofortigem Speichern per AJAX.
* Suchen & Ersetzen-Regeln, eigene Felder, Formel-Felder (inkl. Excel-Formeln) und Datumsformatierung pro Feld.
* Erstelldatum als auswählbare Spalte, verbesserte Export-Regeln und automatische Aktualisierung der Vorschau.
* Neue Übersetzungen (EN, ES, FR, IT, FI, EL, SV, DA, PL) und Textdomain-Dateinamen korrigiert.

= Version 1.0.1 – 19. Dezember 2025 =
* Neue Bereinigungsoption für Elementor Submissions: Nur wenn die Warnung aktiviert ist, werden alte Einträge nach dem gewählten Stunden-/Tage-/Monate-Intervall gelöscht.

= Version 1.0.1 – 19. Dezember 2025 =
* Neue Standardfarbpalette, die auf dem aktuellen Design basiert, und das Chart-Footer-Link-Layout passt sich jetzt exakt der Breite von Diagramm/Tabelle an.

= Version 1.0.1 – 19. Dezember 2025 =
* Neue Option im E-Mail Versand: HTML-Export kann weiterhin gespeichert werden, muss aber nicht mehr zwangsläufig als Anhang versendet werden.

= Version 1.0.1 - 26. August 2025 =
* Filtern über Checkboxen statt Menü. So kann man einzelne Formulare aus der Statistik ausblenden. 


= Version 1.0.1 - 28. April 2025 =
* Kurve geht nur bis zum aktuellen Datum


= Version 1.0.1 - 25. März 2025 =
* Farbschema und Transparenzen im Kurvendiagramm geändert

= Version 1.0.1 - 1. September 2024 =
* Hinzugefügt: Grafik filter nach Formular
*/

function mdp_efs_get_missing_dependencies() {
    if (!function_exists('is_plugin_active')) {
        require_once ABSPATH . 'wp-admin/includes/plugin.php';
    }
    $missing = array();
    if (!is_plugin_active('elementor/elementor.php')) {
        $missing[] = 'Elementor';
    }
    if (!is_plugin_active('elementor-pro/elementor-pro.php')) {
        $missing[] = 'Elementor Pro';
    }
    return $missing;
}

function mdp_efs_missing_dependencies_notice() {
    $missing = mdp_efs_get_missing_dependencies();
    if (empty($missing)) {
        return;
    }
    echo '<div class="notice notice-error"><p>' .
        esc_html(sprintf(
            __('Elementor Forms Statistics benötigt: %s. Bitte Plugin installieren/aktivieren.', 'elementor-forms-statistics'),
            implode(', ', $missing)
        )) .
        '</p></div>';
}

add_action('admin_notices', 'mdp_efs_missing_dependencies_notice');

if (!defined('MDP_ARCHIVE_TABLE_SCHEMA_VERSION')) {
    define('MDP_ARCHIVE_TABLE_SCHEMA_VERSION', 3);
}

function mdp_get_archive_table_name() {
    global $wpdb;
    return $wpdb->prefix . 'mdp_form_stats';
}

function mdp_install_archive_table() {
    global $wpdb;
    $table_name = mdp_get_archive_table_name();
    $current_version = (int) get_option('mdp_efs_archive_schema_version', 0);
    if (mdp_archive_table_exists() && $current_version === MDP_ARCHIVE_TABLE_SCHEMA_VERSION) {
        return;
    }
    $charset_collate = $wpdb->get_charset_collate();
    $sql = "CREATE TABLE {$table_name} (
        id bigint(20) unsigned NOT NULL AUTO_INCREMENT,
        form_id varchar(190) NOT NULL,
        form_name varchar(190) NOT NULL DEFAULT '',
        referer_title text,
        form_title text,
        year smallint(4) NOT NULL,
        month tinyint(2) NOT NULL,
        total int(11) NOT NULL DEFAULT 0,
        updated_at datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        PRIMARY KEY  (id),
        UNIQUE KEY form_month_name (form_id(150),form_name(150),referer_title(150),year,month)
    ) {$charset_collate};";

    require_once ABSPATH . 'wp-admin/includes/upgrade.php';
    dbDelta($sql);

    if (mdp_archive_table_exists() && $current_version < MDP_ARCHIVE_TABLE_SCHEMA_VERSION) {
        $column = $wpdb->get_var("SHOW COLUMNS FROM {$table_name} LIKE 'form_name'");
        if (!$column) {
            $wpdb->query("ALTER TABLE {$table_name} ADD COLUMN form_name varchar(190) NOT NULL DEFAULT '' AFTER form_id");
        }
        $referer_column = $wpdb->get_var("SHOW COLUMNS FROM {$table_name} LIKE 'referer_title'");
        if (!$referer_column) {
            $wpdb->query("ALTER TABLE {$table_name} ADD COLUMN referer_title text AFTER form_name");
        }
        $wpdb->query("UPDATE {$table_name} SET form_name = form_title WHERE form_name = '' OR form_name IS NULL");

        $indexes = $wpdb->get_results("SHOW INDEX FROM {$table_name}", ARRAY_A);
        $index_names = array();
        foreach ($indexes as $index) {
            if (!empty($index['Key_name'])) {
                $index_names[$index['Key_name']] = true;
            }
        }
        if (isset($index_names['form_month'])) {
            $wpdb->query("ALTER TABLE {$table_name} DROP INDEX form_month");
        }
        if (isset($index_names['form_month_name'])) {
            $wpdb->query("ALTER TABLE {$table_name} DROP INDEX form_month_name");
        }
        if (!isset($index_names['form_month_name'])) {
            $wpdb->query("ALTER TABLE {$table_name} ADD UNIQUE KEY form_month_name (form_id(150),form_name(150),referer_title(150),year,month)");
        }

        update_option('mdp_efs_archive_schema_version', MDP_ARCHIVE_TABLE_SCHEMA_VERSION);
    }

    if (mdp_archive_table_exists()) {
        update_option('mdp_efs_archive_schema_version', MDP_ARCHIVE_TABLE_SCHEMA_VERSION);
    }
}

register_uninstall_hook(__FILE__, 'mdp_uninstall_plugin');

function mdp_activate_plugin() {
    mdp_install_archive_table();
    update_option('mdp_efs_show_archive_notice', 1);
    mdp_schedule_submission_cleanup(true);
}
register_activation_hook(__FILE__, 'mdp_activate_plugin');
add_action('plugins_loaded', 'mdp_install_archive_table');
add_action('init', 'mdp_archive_maybe_sync');
add_action('admin_post_mdp_run_archive_import', 'mdp_handle_archive_import');

function mdp_archive_table_exists() {
    global $wpdb;
    $table_name = mdp_get_archive_table_name();
    $like = $wpdb->prepare("SHOW TABLES LIKE %s", $table_name);
    return $wpdb->get_var($like) === $table_name;
}

function mdp_archive_has_data() {
    if (!mdp_archive_table_exists()) {
        return false;
    }
    global $wpdb;
    $table = mdp_get_archive_table_name();
    $count = (int) $wpdb->get_var("SELECT COUNT(*) FROM {$table} LIMIT 1");
    return $count > 0;
}

function mdp_should_use_archive() {
    if (!mdp_archive_table_exists()) {
        return false;
    }
    if (get_option('mdp_efs_archive_initialized')) {
        return true;
    }
    if (mdp_archive_has_data()) {
        update_option('mdp_efs_archive_initialized', 1);
        return true;
    }
    return false;
}

function mdp_get_archived_form_titles() {
    global $wpdb;
    if (!mdp_archive_table_exists()) {
        return [];
    }
    $table = mdp_get_archive_table_name();
    return $wpdb->get_results("
        SELECT form_id,
               MAX(NULLIF(form_name, '')) AS form_title
        FROM {$table}
        GROUP BY form_id
    ");
}

function mdp_get_elementor_form_snapshot_titles() {
    if (!class_exists('ElementorPro\\Modules\\Forms\\Submissions\\Database\\Repositories\\Form_Snapshot_Repository')) {
        return [];
    }

    try {
        $repository = ElementorPro\Modules\Forms\Submissions\Database\Repositories\Form_Snapshot_Repository::instance();
    } catch (Throwable $e) {
        return [];
    }

    if (!$repository) {
        return [];
    }

    $titles = [];
    $snapshots = $repository->all();
    if (empty($snapshots)) {
        return $titles;
    }

    foreach ($snapshots as $snapshot) {
        if (empty($snapshot->id) || empty($snapshot->name)) {
            continue;
        }
        $titles[$snapshot->id] = sanitize_text_field($snapshot->name);
    }
    return $titles;
}

function mdp_archive_sync_new_entries($force_full = false) {
    global $wpdb;
    mdp_install_archive_table();
    $archive_table = mdp_get_archive_table_name();
    $source_table = $wpdb->prefix . 'e_submissions';
    $last_id = $force_full ? 0 : (int) get_option('mdp_efs_archive_last_id', 0);

    if ($force_full) {
        $wpdb->query("TRUNCATE TABLE {$archive_table}");
    }

    $where = "s.status NOT LIKE '%trash%'";
    if ($last_id > 0) {
        $where .= $wpdb->prepare(" AND s.ID > %d", $last_id);
    }
    $where .= mdp_get_email_exclusion_clause('s');

    $sql = "
        SELECT s.element_id AS form_id,
               NULLIF(s.form_name, '') AS form_name,
               NULLIF(s.referer_title, '') AS referer_title,
               MAX(NULLIF(s.form_name, '')) AS form_title,
               YEAR(s.created_at_gmt) AS year_value,
               MONTH(s.created_at_gmt) AS month_value,
               COUNT(*) AS total
        FROM {$source_table} s
        WHERE {$where}
        GROUP BY s.element_id, form_name, referer_title, year_value, month_value
    ";
    $rows = $wpdb->get_results($sql);

    if ($rows) {
        foreach ($rows as $row) {
            if (empty($row->form_id)) {
                continue;
            }
            $form_title = $row->form_title ? sanitize_text_field($row->form_title) : '';
            $form_name = isset($row->form_name) ? sanitize_text_field($row->form_name) : '';
            $referer_title = isset($row->referer_title) ? sanitize_text_field($row->referer_title) : '';
            $wpdb->query($wpdb->prepare(
                "INSERT INTO {$archive_table} (form_id, form_name, referer_title, form_title, year, month, total)
                 VALUES (%s, %s, %s, %s, %d, %d, %d)
                 ON DUPLICATE KEY UPDATE total = total + VALUES(total),
                   form_title = CASE WHEN VALUES(form_title) = '' THEN form_title ELSE VALUES(form_title) END,
                   form_name = CASE WHEN VALUES(form_name) = '' THEN form_name ELSE VALUES(form_name) END,
                   referer_title = CASE WHEN VALUES(referer_title) = '' THEN referer_title ELSE VALUES(referer_title) END",
                $row->form_id,
                $form_name,
                $referer_title,
                $form_title,
                (int) $row->year_value,
                (int) $row->month_value,
                (int) $row->total
            ));
        }
        $max_id_sql = "SELECT MAX(ID) FROM {$source_table} WHERE status NOT LIKE '%trash%'";
        if ($last_id > 0) {
            $max_id_sql .= $wpdb->prepare(" AND ID > %d", $last_id);
        }
        $max_id = (int) $wpdb->get_var($max_id_sql);
        if ($force_full && !$max_id) {
            $max_id = (int) $wpdb->get_var("SELECT MAX(ID) FROM {$source_table}");
        }
        if ($max_id) {
            update_option('mdp_efs_archive_last_id', max($max_id, $last_id));
        }
        update_option('mdp_efs_archive_last_run', current_time('mysql'));
        update_option('mdp_efs_archive_initialized', 1);
    } elseif ($force_full) {
        update_option('mdp_efs_archive_last_run', current_time('mysql'));
        update_option('mdp_efs_archive_initialized', 1);
    }

    update_option('mdp_efs_archive_last_sync', time());
    return count($rows);
}

function mdp_archive_maybe_sync() {
    if (!get_option('mdp_efs_archive_initialized')) {
        return;
    }
    $last_sync = (int) get_option('mdp_efs_archive_last_sync', 0);
    if ($last_sync && (time() - $last_sync) < HOUR_IN_SECONDS) {
        return;
    }
    mdp_archive_sync_new_entries(false);
}

function mdp_handle_archive_import() {
    if (!mdp_user_can_access_menu('statistiken-archiv')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }
    if (!isset($_POST['mdp_archive_nonce']) || !wp_verify_nonce($_POST['mdp_archive_nonce'], 'mdp_archive_import')) {
        wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
    }
    $initialized = (bool) get_option('mdp_efs_archive_initialized');
    $force_full = !$initialized || !empty($_POST['mdp_archive_reset']);
    mdp_archive_sync_new_entries($force_full);
    $redirect = wp_get_referer();
    if (!$redirect) {
        $redirect = admin_url('admin.php?page=statistiken-archiv');
    }
    $status = $force_full ? 'initialized' : 'synced';
    wp_safe_redirect(add_query_arg('mdp_archive_status', $status, $redirect));
    exit;
}

/* CSS */
function enqueue_custom_plugin_styles() {
    wp_enqueue_style('custom-plugin-styles', plugin_dir_url(__FILE__) . 'assets/css/style.css');
}
add_action('admin_enqueue_scripts', 'enqueue_custom_plugin_styles');

function mdp_get_menu_role_choices() {
    $wp_roles = wp_roles();
    if (!$wp_roles || empty($wp_roles->roles)) {
        return array();
    }
    $choices = array();
    foreach ($wp_roles->roles as $role_slug => $role_data) {
        $choices[$role_slug] = isset($role_data['name']) ? $role_data['name'] : $role_slug;
    }
    return $choices;
}

function mdp_get_menu_items() {
    return array(
        'statistiken' => __('Statistik (Hauptmenü)', 'elementor-forms-statistics'),
        'statistiken-einstellungen' => __('Einstellungen', 'elementor-forms-statistics'),
        'statistiken-emailversand' => __('E-Mail Versand', 'elementor-forms-statistics'),
        'statistiken-archiv' => __('Archiv', 'elementor-forms-statistics'),
        'statistiken-export' => __('Export', 'elementor-forms-statistics'),
        'elementor-submissions' => __('Anfragen', 'elementor-forms-statistics'),
    );
}

function mdp_get_default_menu_roles() {
    $choices = mdp_get_menu_role_choices();
    $defaults = array();
    $edit_posts_roles = array();
    foreach ($choices as $role_slug => $role_label) {
        $role_obj = get_role($role_slug);
        if ($role_obj && $role_obj->has_cap('edit_posts')) {
            $edit_posts_roles[] = $role_slug;
        }
    }
    $defaults['statistiken'] = $edit_posts_roles;
    $defaults['statistiken-einstellungen'] = $edit_posts_roles;
    $defaults['statistiken-emailversand'] = $edit_posts_roles;
    $defaults['statistiken-archiv'] = $edit_posts_roles;
    $defaults['statistiken-export'] = array('administrator');
    $defaults['elementor-submissions'] = array('editor');
    return $defaults;
}

function mdp_get_menu_roles() {
    $stored = get_option('mdp_efs_menu_roles', null);
    $defaults = mdp_get_default_menu_roles();
    $choices = mdp_get_menu_role_choices();
    $valid_roles = array_keys($choices);
    $menu_items = array_keys(mdp_get_menu_items());

    if ($stored === null) {
        $stored = $defaults;
        $legacy_export = get_option('mdp_efs_export_roles', null);
        if (is_array($legacy_export)) {
            $legacy_allowed = array();
            foreach ($legacy_export as $role) {
                $role = sanitize_key($role);
                if ($role !== '' && in_array($role, $valid_roles, true)) {
                    $legacy_allowed[] = $role;
                }
            }
            if (!empty($legacy_allowed)) {
                $stored['statistiken-export'] = array_values(array_unique($legacy_allowed));
            }
        }
        update_option('mdp_efs_menu_roles', $stored);
    }

    if (!is_array($stored)) {
        $stored = array();
    }

    $normalized = array();
    foreach ($menu_items as $menu_key) {
        $normalized[$menu_key] = array();
        $roles = isset($stored[$menu_key]) && is_array($stored[$menu_key]) ? $stored[$menu_key] : $defaults[$menu_key];
        foreach ($roles as $role) {
            $role = sanitize_key($role);
            if ($role !== '' && in_array($role, $valid_roles, true)) {
                $normalized[$menu_key][] = $role;
            }
        }
        $normalized[$menu_key] = array_values(array_unique($normalized[$menu_key]));
    }

    return $normalized;
}

function mdp_user_can_access_menu($menu_key, $user = null) {
    if (!is_user_logged_in()) {
        return false;
    }
    $menu_roles = mdp_get_menu_roles();
    if (empty($menu_roles[$menu_key])) {
        return false;
    }
    $user = $user instanceof WP_User ? $user : wp_get_current_user();
    if (!$user || empty($user->roles)) {
        return false;
    }
    foreach ($user->roles as $role) {
        if (in_array($role, $menu_roles[$menu_key], true)) {
            return true;
        }
    }
    return false;
}

function mdp_user_can_access_export_menu($user = null) {
    return mdp_user_can_access_menu('statistiken-export', $user);
}

require_once __DIR__ . '/elementor-form-submissions-access.php';

function mdp_get_export_fields_option() {
    $stored = get_option('mdp_efs_export_fields', []);
    return is_array($stored) ? $stored : [];
}

function mdp_get_export_replace_rules_option() {
    $stored = get_option('mdp_efs_export_replace_rules', []);
    return is_array($stored) ? $stored : [];
}

function mdp_get_export_date_candidate_keys($form_id) {
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return [];
    }
    global $wpdb;
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $regex = '^[0-9]{4}-[0-9]{2}-[0-9]{2}';
    $rows = $wpdb->get_results(
        $wpdb->prepare(
            "SELECT DISTINCT ev.`key`
             FROM {$values_table} ev
             INNER JOIN {$submissions_table} s ON s.ID = ev.submission_id
             WHERE s.element_id = %s
               AND s.status NOT LIKE %s
               AND ev.`value` REGEXP %s
             LIMIT 100",
            $form_id,
            '%trash%',
            $regex
        ),
        ARRAY_A
    );
    $keys = [];
    foreach ($rows as $row) {
        $key = isset($row['key']) ? sanitize_text_field($row['key']) : '';
        if ($key !== '') {
            $keys[] = $key;
        }
    }
    if (!in_array('created_at', $keys, true)) {
        $keys[] = 'created_at';
    }
    return $keys;
}

function mdp_get_export_date_format_choices() {
    return array(
        '' => __('Original', 'elementor-forms-statistics'),
        'd.m.Y' => __('TT.MM.JJJJ', 'elementor-forms-statistics'),
    );
}

function mdp_get_export_formula_fields_option() {
    $stored = get_option('mdp_efs_export_formula_fields', []);
    return is_array($stored) ? $stored : [];
}

function mdp_get_export_formula_fields($form_id) {
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return [];
    }
    $stored = mdp_get_export_formula_fields_option();
    $fields = isset($stored[$form_id]) && is_array($stored[$form_id]) ? $stored[$form_id] : [];
    $clean = [];
    foreach ($fields as $field) {
        if (!is_array($field)) {
            continue;
        }
        $key = isset($field['key']) ? sanitize_text_field($field['key']) : '';
        $label = isset($field['label']) ? sanitize_text_field($field['label']) : '';
        $formula = isset($field['formula']) ? sanitize_text_field($field['formula']) : '';
        if ($key === '' || $formula === '') {
            continue;
        }
        if ($label === '') {
            $label = $key;
        }
        $clean[] = array(
            'key' => $key,
            'label' => $label,
            'formula' => $formula,
        );
    }
    return $clean;
}

function mdp_get_export_replace_rules($form_id) {
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return [];
    }
    $stored = mdp_get_export_replace_rules_option();
    $rules = isset($stored[$form_id]) && is_array($stored[$form_id]) ? $stored[$form_id] : [];
    $clean = [];
    foreach ($rules as $rule) {
        if (!is_array($rule)) {
            continue;
        }
        $find = isset($rule['find']) ? (string) $rule['find'] : '';
        $replace = isset($rule['replace']) ? (string) $rule['replace'] : '';
        if ($find === '') {
            continue;
        }
        $clean[] = array(
            'find' => $find,
            'replace' => $replace,
        );
    }
    return $clean;
}

function mdp_apply_export_replace_rules($form_id, $value) {
    $rules = mdp_get_export_replace_rules($form_id);
    if (empty($rules)) {
        return $value;
    }
    foreach ($rules as $rule) {
        $find = isset($rule['find']) ? $rule['find'] : '';
        if ($find === '') {
            continue;
        }
        $replace = isset($rule['replace']) ? $rule['replace'] : '';
        $value = str_replace($find, $replace, $value);
    }
    return $value;
}

function mdp_apply_export_date_format($value, $format) {
    if (!is_string($format) || $format === '') {
        return $value;
    }
    if (!is_string($value)) {
        return $value;
    }
    $value = trim($value);
    if ($value === '' || $value[0] === '=') {
        return $value;
    }
    $parts = explode(' | ', $value);
    $formatted = [];
    foreach ($parts as $part) {
        $part = trim($part);
        if ($part === '') {
            $formatted[] = $part;
            continue;
        }
        if (preg_match('/^(\\d{4})-(\\d{2})-(\\d{2})/', $part, $matches)) {
            $timestamp = strtotime($matches[1] . '-' . $matches[2] . '-' . $matches[3] . ' UTC');
            if ($timestamp) {
                $formatted[] = wp_date($format, $timestamp);
                continue;
            }
        }
        $formatted[] = $part;
    }
    return implode(' | ', $formatted);
}

function mdp_split_export_formula_tokens($formula) {
    $tokens = [];
    $current = '';
    $in_quote = false;
    $quote_char = '';
    $length = strlen($formula);
    for ($i = 0; $i < $length; $i++) {
        $char = $formula[$i];
        if ($in_quote) {
            if ($char === $quote_char) {
                $in_quote = false;
            }
            $current .= $char;
            continue;
        }
        if ($char === '"' || $char === "'") {
            $in_quote = true;
            $quote_char = $char;
            $current .= $char;
            continue;
        }
        if ($char === '&') {
            $tokens[] = $current;
            $current = '';
            continue;
        }
        $current .= $char;
    }
    if ($current !== '') {
        $tokens[] = $current;
    }
    return $tokens;
}

function mdp_normalize_formula_key($key) {
    $key = (string) $key;
    $key = strtolower($key);
    $key = str_replace('ß', 'ss', $key);
    $key = preg_replace('/[^a-z0-9_]/', '', $key);
    return $key;
}

function mdp_evaluate_export_formula($formula, $value_map) {
    $formula = is_string($formula) ? trim($formula) : '';
    if ($formula !== '' && $formula[0] === '=') {
        return $formula;
    }
    $tokens = mdp_split_export_formula_tokens($formula);
    if (count($tokens) === 1) {
        $single = trim($tokens[0]);
        if ($single !== '') {
            $first = substr($single, 0, 1);
            $last = substr($single, -1);
            if (!(($first === '"' && $last === '"') || ($first === "'" && $last === "'"))) {
                $tokens = array($single);
            }
        }
    }
    $output = '';
    $normalized_map = [];
    foreach ($value_map as $map_key => $map_value) {
        $normalized_map[mdp_normalize_formula_key($map_key)] = $map_value;
    }
    foreach ($tokens as $token) {
        $token = trim($token);
        if ($token === '') {
            continue;
        }
        $first = substr($token, 0, 1);
        $last = substr($token, -1);
        if (($first === '"' && $last === '"') || ($first === "'" && $last === "'")) {
            $output .= substr($token, 1, -1);
            continue;
        }
        $key = $token;
        if (preg_match('/^[a-zA-Z0-9_]+$/', $key) === 1 && isset($value_map[$key])) {
            $output .= $value_map[$key];
            continue;
        }
        if (isset($value_map[$key])) {
            $output .= $value_map[$key];
            continue;
        }
        $lower = strtolower($key);
        if (isset($value_map[$lower])) {
            $output .= $value_map[$lower];
            continue;
        }
        $normalized = mdp_normalize_formula_key($key);
        $output .= isset($normalized_map[$normalized]) ? $normalized_map[$normalized] : '';
    }
    return $output;
}

function mdp_get_submission_value_map($values_by_submission, $submission_id, $form_id) {
    $map = [];
    if (!isset($values_by_submission[$submission_id])) {
        return $map;
    }
    foreach ($values_by_submission[$submission_id] as $key => $values) {
        if (empty($values)) {
            $map[$key] = '';
            continue;
        }
        $map[$key] = implode(' | ', $values);
    }
    return $map;
}
function mdp_get_last_export_form_id($user_id) {
    $user_id = (int) $user_id;
    if ($user_id <= 0) {
        return '';
    }
    $value = get_user_meta($user_id, 'mdp_efs_last_export_form_id', true);
    return is_string($value) ? sanitize_text_field($value) : '';
}

function mdp_get_last_export_format($user_id) {
    $user_id = (int) $user_id;
    if ($user_id <= 0) {
        return 'csv';
    }
    // CSV-only export: ignore persisted format values.
    return 'csv';
}

function mdp_set_last_export_format($user_id, $format) {
    $user_id = (int) $user_id;
    if ($user_id <= 0) {
        return;
    }
    // Keep meta consistent even if older UI posted "excel".
    update_user_meta($user_id, 'mdp_efs_last_export_format', 'csv');
}

function mdp_format_submission_created_at($created_at_gmt) {
    $created_at_gmt = is_string($created_at_gmt) ? trim($created_at_gmt) : '';
    if ($created_at_gmt === '') {
        return '';
    }
    $timestamp = strtotime($created_at_gmt . ' UTC');
    if (!$timestamp) {
        return $created_at_gmt;
    }
    return wp_date('d.m.Y', $timestamp);
}

function mdp_set_last_export_form_id($user_id, $form_id) {
    $user_id = (int) $user_id;
    if ($user_id <= 0) {
        return;
    }
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return;
    }
    update_user_meta($user_id, 'mdp_efs_last_export_form_id', $form_id);
}


function mdp_set_submission_value($submission_id, $key, $value) {
    global $wpdb;
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submission_id = (int) $submission_id;
    $key = sanitize_key($key);
    if ($submission_id <= 0 || $key === '') {
        return;
    }
    $wpdb->delete($values_table, array(
        'submission_id' => $submission_id,
        'key' => $key,
    ));
    if ($value === '' || $value === null) {
        return;
    }
    $wpdb->insert($values_table, array(
        'submission_id' => $submission_id,
        'key' => $key,
        'value' => $value,
    ));
}

function mdp_delete_submission_value($submission_id, $key) {
    global $wpdb;
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submission_id = (int) $submission_id;
    $key = sanitize_key($key);
    if ($submission_id <= 0 || $key === '') {
        return;
    }
    $wpdb->delete($values_table, array(
        'submission_id' => $submission_id,
        'key' => $key,
    ));
}

function mdp_get_form_field_keys($form_id) {
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return [];
    }
    if (class_exists('ElementorPro\\Modules\\Forms\\Submissions\\Database\\Repositories\\Form_Snapshot_Repository')) {
        try {
            $repository = ElementorPro\Modules\Forms\Submissions\Database\Repositories\Form_Snapshot_Repository::instance();
        } catch (Throwable $e) {
            $repository = null;
        }
        if ($repository) {
            global $wpdb;
            $submissions_table = $wpdb->prefix . 'e_submissions';
            $post_id = (int) $wpdb->get_var($wpdb->prepare(
                "SELECT post_id FROM {$submissions_table}
                 WHERE element_id = %s
                 ORDER BY created_at_gmt DESC
                 LIMIT 1",
                $form_id
            ));
            if ($post_id > 0) {
                $snapshot = $repository->find($post_id, $form_id);
                if ($snapshot && !empty($snapshot->fields) && is_array($snapshot->fields)) {
                    $keys = [];
                    foreach ($snapshot->fields as $field) {
                        if (!is_array($field)) {
                            continue;
                        }
                        $key = isset($field['id']) ? sanitize_text_field($field['id']) : '';
                        if ($key !== '') {
                            $keys[] = $key;
                        }
                    }
                    $keys = array_values(array_unique($keys));
                    if (!empty($keys)) {
                        if (!in_array('created_at', $keys, true)) {
                            $keys[] = 'created_at';
                        }
                        return $keys;
                    }
                }
            }
            $snapshots = $repository->all();
            foreach ($snapshots as $snapshot) {
                if (empty($snapshot->id) || (string) $snapshot->id !== (string) $form_id) {
                    continue;
                }
                $keys = [];
                if (!empty($snapshot->fields) && is_array($snapshot->fields)) {
                    foreach ($snapshot->fields as $field) {
                        if (!is_array($field)) {
                            continue;
                        }
                        $key = isset($field['id']) ? sanitize_text_field($field['id']) : '';
                        if ($key !== '') {
                            $keys[] = $key;
                        }
                    }
                }
                $keys = array_values(array_unique($keys));
                if (!empty($keys)) {
                    if (!in_array('created_at', $keys, true)) {
                        $keys[] = 'created_at';
                    }
                    return $keys;
                }
                break;
            }
        }
    }
    global $wpdb;
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $rows = $wpdb->get_results($wpdb->prepare(
        "SELECT DISTINCT ev.`key`
         FROM {$values_table} ev
         INNER JOIN {$submissions_table} s ON s.ID = ev.submission_id
         WHERE s.element_id = %s
           AND s.status NOT LIKE %s
         ORDER BY ev.`key`",
        $form_id,
        '%trash%'
    ), ARRAY_A);
    $keys = [];
    foreach ($rows as $row) {
        $key = isset($row['key']) ? sanitize_text_field($row['key']) : '';
        if ($key !== '') {
            $keys[] = $key;
        }
    }
    if (!in_array('created_at', $keys, true)) {
        $keys[] = 'created_at';
    }
    return $keys;
}

function mdp_get_export_fields_for_form($form_id) {
    $available_keys = mdp_get_form_field_keys($form_id);
    $formula_fields = mdp_get_export_formula_fields($form_id);
    $formula_map = [];
    foreach ($formula_fields as $formula_entry) {
        if (!empty($formula_entry['key'])) {
            $formula_map[$formula_entry['key']] = $formula_entry;
        }
    }
    $stored_all = mdp_get_export_fields_option();
    $stored = isset($stored_all[$form_id]) && is_array($stored_all[$form_id]) ? $stored_all[$form_id] : [];
    $date_format_choices = mdp_get_export_date_format_choices();
    $allowed_date_formats = array_keys($date_format_choices);
    $ordered = [];
    $seen = [];
    foreach ($stored as $entry) {
        if (!is_array($entry)) {
            continue;
        }
        $key = isset($entry['key']) ? sanitize_text_field($entry['key']) : '';
        $is_custom = !empty($entry['custom']);
        $is_formula = isset($formula_map[$key]);
        if ($key === '' || (!$is_custom && !$is_formula && !in_array($key, $available_keys, true))) {
            continue;
        }
        $label = isset($entry['label']) && $entry['label'] !== '' ? sanitize_text_field($entry['label']) : $key;
        $date_format = isset($entry['date_format']) ? sanitize_text_field($entry['date_format']) : '';
        if (!in_array($date_format, $allowed_date_formats, true)) {
            $date_format = '';
        }
        $include = true;
        if (isset($entry['include'])) {
            $include = (bool) $entry['include'];
        }
        $ordered[] = array(
            'key' => $key,
            'label' => $label,
            'include' => $include,
            'custom' => $is_custom,
            'formula' => $is_formula,
            'date_format' => $date_format,
        );
        $seen[$key] = true;
    }
    if (!empty($available_keys)) {
        foreach ($available_keys as $key) {
            if (isset($seen[$key])) {
                continue;
            }
            $label = $key;
            if ($key === 'created_at') {
                $label = __('Erstelldatum', 'elementor-forms-statistics');
            }
            $ordered[] = array(
                'key' => $key,
                'label' => $label,
                'include' => true,
                'custom' => false,
                'formula' => false,
                'date_format' => '',
            );
        }
    }
    if (!empty($formula_fields)) {
        foreach ($formula_fields as $formula_entry) {
            $key = $formula_entry['key'];
            if ($key === '' || isset($seen[$key])) {
                continue;
            }
            $ordered[] = array(
                'key' => $key,
                'label' => $formula_entry['label'],
                'include' => true,
                'custom' => false,
                'formula' => true,
                'date_format' => '',
            );
            $seen[$key] = true;
        }
    }
    return $ordered;
}

function custom_menu_item() {
    $can_view_stats = mdp_user_can_access_menu('statistiken')
        || mdp_user_can_access_menu('statistiken-einstellungen')
        || mdp_user_can_access_menu('statistiken-emailversand')
        || mdp_user_can_access_menu('statistiken-archiv')
        || mdp_user_can_access_menu('statistiken-export')
        || mdp_user_can_access_menu('elementor-submissions');

    if (!$can_view_stats) {
        return;
    }

    add_menu_page(
        __('Statistik', 'elementor-forms-statistics'), // The title of the menu item
        __('Statistik', 'elementor-forms-statistics'), // The name of the menu item in the navigation bar
        'edit_posts', // Required capability to see the menu item (for editors and admins)
        'statistiken', // A neutral slug for the menu item
        'custom_menu_callback', // The function to call when the menu item is selected
        'dashicons-chart-line', // The icon for the menu item
        2 // The position of the menu item in the navigation bar
    );

    if (mdp_user_can_access_menu('statistiken-einstellungen')) {
        add_submenu_page(
            'statistiken',
            __('Einstellungen', 'elementor-forms-statistics'),
            __('Einstellungen', 'elementor-forms-statistics'),
            'edit_posts',
            'statistiken-einstellungen',
            'mdp_settings_page_callback'
        );
    }

    if (mdp_user_can_access_menu('statistiken-emailversand')) {
        add_submenu_page(
            'statistiken',
            __('E-Mail Versand', 'elementor-forms-statistics'),
            __('E-Mail Versand', 'elementor-forms-statistics'),
            'edit_posts',
            'statistiken-emailversand',
            'mdp_email_settings_page_callback'
        );
    }

    if (mdp_user_can_access_menu('statistiken-archiv')) {
        add_submenu_page(
            'statistiken',
            __('Archiv', 'elementor-forms-statistics'),
            __('Archiv', 'elementor-forms-statistics'),
            'edit_posts',
            'statistiken-archiv',
            'mdp_archive_page_callback'
        );
    }

    if (mdp_user_can_access_export_menu()) {
        add_submenu_page(
            'statistiken',
            __('Export', 'elementor-forms-statistics'),
            __('Export', 'elementor-forms-statistics'),
            'edit_posts',
            'statistiken-export',
            'mdp_export_page_callback'
        );
    }

}

function mdp_archive_setup_notice() {
    if (!mdp_user_can_access_menu('statistiken-archiv')) {
        return;
    }
    $show = get_option('mdp_efs_show_archive_notice', 0);
    if (get_option('mdp_efs_archive_initialized')) {
        return;
    }
    if (!$show) {
        return;
    }
    $settings_url = admin_url('admin.php?page=statistiken-archiv');
    $message = '<strong>' . __('Elementor Forms Statistics.', 'elementor-forms-statistics') . '</strong> ';
    $message .= __('Bitte Plugin einmalig initialisieren, damit die Statistikdaten gesichert und anschließend automatisch synchronisiert werden.', 'elementor-forms-statistics') . ' ';
    $message .= sprintf(
        __('Öffne die %s, um den Vorgang manuell zu starten und die Sicherung zu prüfen.', 'elementor-forms-statistics'),
        '<a href="' . esc_url($settings_url) . '">' . esc_html__('Einstellungen zur Statistik-Sicherung', 'elementor-forms-statistics') . '</a>'
    );
    $dismiss_url = wp_nonce_url(add_query_arg('mdp_archive_notice_dismiss', '1', admin_url('admin.php')), 'mdp_archive_notice_dismiss');
    echo '<div class="notice notice-info">';
    echo '<p>' . $message . '</p>';
    echo '<p><a href="' . esc_url($dismiss_url) . '">' . esc_html__('Hinweis ausblenden', 'elementor-forms-statistics') . '</a></p>';
    echo '</div>';
}

function mdp_handle_archive_notice_dismiss() {
    if (!mdp_user_can_access_menu('statistiken-archiv')) {
        return;
    }
    if (empty($_GET['mdp_archive_notice_dismiss'])) {
        return;
    }
    if (empty($_GET['_wpnonce']) || !wp_verify_nonce(wp_unslash($_GET['_wpnonce']), 'mdp_archive_notice_dismiss')) {
        return;
    }
    update_option('mdp_efs_show_archive_notice', 0);
}
add_action('admin_notices', 'mdp_archive_setup_notice');
add_action('admin_init', 'mdp_handle_archive_notice_dismiss');
add_action('admin_menu', 'custom_menu_item');
add_filter('plugin_action_links_' . plugin_basename(__FILE__), 'mdp_efs_plugin_action_links');
add_filter('plugin_row_meta', 'mdp_efs_plugin_row_meta', 10, 2);

function mdp_efs_plugin_action_links($links) {
    if (mdp_user_can_access_menu('statistiken-einstellungen')) {
        $settings_url = admin_url('admin.php?page=statistiken-einstellungen');
        $links[] = '<a href="' . esc_url($settings_url) . '">' . esc_html__('Einstellungen', 'elementor-forms-statistics') . '</a>';
    }
    return $links;
}

function mdp_efs_plugin_row_meta($links, $file) {
    if ($file !== plugin_basename(__FILE__)) {
        return $links;
    }
    $links[] = '<span class="mdp-efs-deps">' .
        esc_html__('Benötigt:', 'elementor-forms-statistics') .
        ' ' .
        esc_html__('Elementor, Elementor Pro', 'elementor-forms-statistics') .
        '</span>';
    return $links;
}
add_action('admin_post_mdp_export_stats_html', 'mdp_export_stats_html');
add_action('admin_post_mdp_export_csv', 'mdp_export_csv');
add_action('admin_post_mdp_save_export_fields', 'mdp_save_export_fields');
add_action('wp_ajax_mdp_get_export_fields', 'mdp_get_export_fields_ajax');
add_action('wp_ajax_mdp_save_export_fields_ajax', 'mdp_save_export_fields_ajax');
add_action('wp_ajax_mdp_save_export_rules_ajax', 'mdp_save_export_rules_ajax');
add_action('wp_ajax_mdp_get_export_rules', 'mdp_get_export_rules_ajax');
add_action('wp_ajax_mdp_save_export_formulas_ajax', 'mdp_save_export_formulas_ajax');
add_action('wp_ajax_mdp_get_export_formulas', 'mdp_get_export_formulas_ajax');
add_action('wp_ajax_mdp_get_export_preview', 'mdp_get_export_preview_ajax');
add_action('wp_ajax_mdp_save_export_last_form', 'mdp_save_export_last_form_ajax');
add_action('wp_ajax_mdp_save_export_last_format', 'mdp_save_export_last_format_ajax');
add_action('init', 'mdp_schedule_submission_cleanup');
add_action('admin_post_mdp_send_stats_now', 'mdp_send_stats_now_handler');
add_filter('cron_schedules', 'mdp_add_custom_cron_schedules');
add_action('init', 'mdp_maybe_schedule_stats_email');
add_action('mdp_send_stats_email', 'mdp_send_stats_email_callback');

function mdp_export_stats_html() {
    if (!mdp_user_can_access_menu('statistiken-emailversand')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }

    $is_inline_request = isset($_REQUEST['inline']) && $_REQUEST['inline'];
    if ($is_inline_request) {
        $nonce = isset($_REQUEST['mdp_inline_nonce']) ? sanitize_text_field($_REQUEST['mdp_inline_nonce']) : '';
        if (!wp_verify_nonce($nonce, 'mdp_export_html_inline')) {
            wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
        }
    } else {
        if (!isset($_POST['mdp_export_nonce']) || !wp_verify_nonce($_POST['mdp_export_nonce'], 'mdp_export_html')) {
            wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
        }
    }

    $html = custom_menu_callback(true, array(
        'export' => true,
        'inline_styles' => true,
        'include_html_document' => true,
        'include_chartjs' => true,
    ));

    $filename = 'anfragen-statistik-' . date('Y-m-d') . '.html';
    nocache_headers();
    $disposition = (isset($_REQUEST['inline']) && $_REQUEST['inline']) ? 'inline' : 'attachment';
    header('Content-Type: text/html; charset=' . get_option('blog_charset'));
    header('Content-Disposition: ' . $disposition . '; filename="' . $filename . '"');
    header('Content-Length: ' . strlen($html));
    echo $html;
    exit;
}

function mdp_normalize_export_value($value) {
    $value = maybe_unserialize($value);
    if (is_array($value)) {
        $flat = [];
        foreach ($value as $item) {
            if (is_scalar($item) || $item === null) {
                $flat[] = (string) $item;
            } else {
                $flat[] = wp_json_encode($item);
            }
        }
        $value = implode(', ', $flat);
    } elseif (is_object($value)) {
        $value = wp_json_encode($value);
    } else {
        $value = (string) $value;
    }
    $value = str_replace(["\r", "\n"], ' ', $value);
    return trim($value);
}

function mdp_export_csv() {
    if (!mdp_user_can_access_export_menu()) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }
    if (!isset($_POST['mdp_export_csv_nonce']) || !wp_verify_nonce($_POST['mdp_export_csv_nonce'], 'mdp_export_csv')) {
        wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
    }
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    // Export format locked to CSV (no Excel variant).
    $format = 'csv';
    mdp_set_last_export_format(get_current_user_id(), $format);
    if ($form_id === '') {
        wp_die(__('Kein Formular ausgewählt.', 'elementor-forms-statistics'));
    }

    $fields = array_values(array_filter(mdp_get_export_fields_for_form($form_id), function($field) {
        return !empty($field['include']);
    }));
    $formula_fields = mdp_get_export_formula_fields($form_id);
    $formula_map = [];
    foreach ($formula_fields as $formula_entry) {
        if (!empty($formula_entry['key'])) {
            $formula_map[$formula_entry['key']] = $formula_entry['formula'];
        }
    }
    $headers = [];
    foreach ($fields as $field) {
        $headers[] = isset($field['label']) && $field['label'] !== '' ? $field['label'] : $field['key'];
    }

    $extension = 'csv';
    $filename = 'elementor-form-' . sanitize_file_name($form_id) . '-' . gmdate('Y-m-d') . '.' . $extension;
    header('Content-Type: text/csv; charset=utf-8');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Pragma: no-cache');
    header('Expires: 0');

    $out = fopen('php://output', 'w');
    if ($out === false) {
        wp_die(__('CSV-Ausgabe fehlgeschlagen.', 'elementor-forms-statistics'));
    }
    fputcsv($out, $headers, ';');

    if (empty($fields)) {
        fclose($out);
        exit;
    }

    global $wpdb;
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submission_rows = $wpdb->get_results($wpdb->prepare(
        "SELECT ID, created_at_gmt
         FROM {$submissions_table}
         WHERE element_id = %s
           AND status NOT LIKE %s
         ORDER BY created_at_gmt DESC",
        $form_id,
        '%trash%'
    ), ARRAY_A);

    if (empty($submission_rows)) {
        fclose($out);
        exit;
    }

    $submission_ids = array_map('intval', wp_list_pluck($submission_rows, 'ID'));
    $created_at_map = [];
    foreach ($submission_rows as $submission_row) {
        $submission_id = isset($submission_row['ID']) ? (int) $submission_row['ID'] : 0;
        if ($submission_id <= 0) {
            continue;
        }
        $created_at_map[$submission_id] = mdp_format_submission_created_at($submission_row['created_at_gmt'] ?? '');
    }
    $values_by_submission = [];
    $chunk_size = 500;
    $field_keys = wp_list_pluck($fields, 'key');
    $allowed_keys = array_unique(array_merge($field_keys, (array) mdp_get_form_field_keys($form_id)));
    foreach (array_chunk($submission_ids, $chunk_size) as $chunk) {
        $placeholders = implode(',', array_fill(0, count($chunk), '%d'));
        $rows = $wpdb->get_results(
            $wpdb->prepare(
                "SELECT submission_id, `key`, `value`
                 FROM {$values_table}
                 WHERE submission_id IN ({$placeholders})",
                $chunk
            ),
            ARRAY_A
        );
        foreach ($rows as $row) {
            $submission_id = isset($row['submission_id']) ? (int) $row['submission_id'] : 0;
            $key = isset($row['key']) ? sanitize_text_field($row['key']) : '';
            if ($submission_id <= 0 || $key === '' || !in_array($key, $allowed_keys, true)) {
                continue;
            }
            $value = mdp_normalize_export_value($row['value']);
            if (!isset($values_by_submission[$submission_id])) {
                $values_by_submission[$submission_id] = [];
            }
            if (!isset($values_by_submission[$submission_id][$key])) {
                $values_by_submission[$submission_id][$key] = [];
            }
            if ($value !== '') {
                $values_by_submission[$submission_id][$key][] = $value;
            }
        }
    }

    foreach ($submission_ids as $submission_id) {
        if (!isset($values_by_submission[$submission_id])) {
            $values_by_submission[$submission_id] = [];
        }
        if (!isset($values_by_submission[$submission_id]['created_at']) && isset($created_at_map[$submission_id])) {
            $values_by_submission[$submission_id]['created_at'] = array($created_at_map[$submission_id]);
        }
        $value_map = mdp_get_submission_value_map($values_by_submission, $submission_id, $form_id);
        $row = [];
        foreach ($fields as $field) {
            $key = $field['key'];
            if (!empty($field['formula']) && isset($formula_map[$key])) {
                $row[] = mdp_apply_export_replace_rules($form_id, mdp_evaluate_export_formula($formula_map[$key], $value_map));
                continue;
            }
            $value = isset($value_map[$key]) ? $value_map[$key] : '';
            if (!empty($field['date_format'])) {
                $value = mdp_apply_export_date_format($value, $field['date_format']);
            }
            $row[] = mdp_apply_export_replace_rules($form_id, $value);
        }
        fputcsv($out, $row, ';');
    }
    fclose($out);
    exit;
}

function mdp_save_export_fields() {
    if (!mdp_user_can_access_export_menu()) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }
    if (!isset($_POST['mdp_export_fields_nonce']) || !wp_verify_nonce($_POST['mdp_export_fields_nonce'], 'mdp_save_export_fields')) {
        wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
    }
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    $payload_raw = isset($_POST['fields_payload']) ? wp_unslash($_POST['fields_payload']) : '';
    if ($form_id === '') {
        wp_die(__('Kein Formular ausgewählt.', 'elementor-forms-statistics'));
    }
    $payload = json_decode($payload_raw, true);
    $payload = is_array($payload) ? $payload : [];
    $available_keys = mdp_get_form_field_keys($form_id);
    $formula_fields = mdp_get_export_formula_fields($form_id);
    $formula_keys = array_values(array_filter(array_map(function($entry) {
        return !empty($entry['key']) ? sanitize_text_field($entry['key']) : '';
    }, $formula_fields)));
    $date_format_choices = mdp_get_export_date_format_choices();
    $allowed_date_formats = array_keys($date_format_choices);
    $valid = [];
    foreach ($payload as $entry) {
        if (!is_array($entry)) {
            continue;
        }
        $key = isset($entry['key']) ? sanitize_text_field($entry['key']) : '';
        $is_custom = !empty($entry['custom']);
        $is_formula = !empty($entry['formula']) && in_array($key, $formula_keys, true);
        if ($key === '' || (!$is_custom && !$is_formula && !in_array($key, $available_keys, true))) {
            continue;
        }
        $label = isset($entry['label']) && $entry['label'] !== '' ? sanitize_text_field($entry['label']) : $key;
        $include = isset($entry['include']) ? (bool) $entry['include'] : true;
        $date_format = isset($entry['date_format']) ? sanitize_text_field($entry['date_format']) : '';
        if (!in_array($date_format, $allowed_date_formats, true)) {
            $date_format = '';
        }
        $valid[] = array(
            'key' => $key,
            'label' => $label,
            'include' => $include,
            'custom' => $is_custom,
            'formula' => $is_formula,
            'date_format' => $date_format,
        );
    }
    $stored = mdp_get_export_fields_option();
    $stored[$form_id] = $valid;
    update_option('mdp_efs_export_fields', $stored);

    $redirect = add_query_arg(
        array(
            'page' => 'statistiken-export',
            'form_id' => rawurlencode($form_id),
            'mdp_export_saved' => 1,
        ),
        admin_url('admin.php')
    );
    wp_safe_redirect($redirect);
    exit;
}

function mdp_get_export_fields_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $fields = mdp_get_export_fields_for_form($form_id);
    $date_candidates = mdp_get_export_date_candidate_keys($form_id);
    wp_send_json_success(array(
        'fields' => $fields,
        'date_candidates' => $date_candidates,
    ));
}

function mdp_save_export_fields_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    $payload_raw = isset($_POST['fields_payload']) ? wp_unslash($_POST['fields_payload']) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $payload = json_decode($payload_raw, true);
    $payload = is_array($payload) ? $payload : [];
    $available_keys = mdp_get_form_field_keys($form_id);
    $formula_fields = mdp_get_export_formula_fields($form_id);
    $formula_keys = array_values(array_filter(array_map(function($entry) {
        return !empty($entry['key']) ? sanitize_text_field($entry['key']) : '';
    }, $formula_fields)));
    $date_format_choices = mdp_get_export_date_format_choices();
    $allowed_date_formats = array_keys($date_format_choices);
    $valid = [];
    foreach ($payload as $entry) {
        if (!is_array($entry)) {
            continue;
        }
        $key = isset($entry['key']) ? sanitize_text_field($entry['key']) : '';
        $is_custom = !empty($entry['custom']);
        $is_formula = !empty($entry['formula']) && in_array($key, $formula_keys, true);
        if ($key === '' || (!$is_custom && !$is_formula && !in_array($key, $available_keys, true))) {
            continue;
        }
        $label = isset($entry['label']) && $entry['label'] !== '' ? sanitize_text_field($entry['label']) : $key;
        $include = isset($entry['include']) ? (bool) $entry['include'] : true;
        $date_format = isset($entry['date_format']) ? sanitize_text_field($entry['date_format']) : '';
        if (!in_array($date_format, $allowed_date_formats, true)) {
            $date_format = '';
        }
        $valid[] = array(
            'key' => $key,
            'label' => $label,
            'include' => $include,
            'custom' => $is_custom,
            'formula' => $is_formula,
            'date_format' => $date_format,
        );
    }
    $stored = mdp_get_export_fields_option();
    $stored[$form_id] = $valid;
    update_option('mdp_efs_export_fields', $stored);
    wp_send_json_success();
}

function mdp_save_export_rules_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    $payload_raw = isset($_POST['rules_payload']) ? wp_unslash($_POST['rules_payload']) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $payload = json_decode($payload_raw, true);
    $payload = is_array($payload) ? $payload : [];
    $valid = [];
    foreach ($payload as $entry) {
        if (!is_array($entry)) {
            continue;
        }
        $find = isset($entry['find']) ? (string) $entry['find'] : '';
        $replace = isset($entry['replace']) ? (string) $entry['replace'] : '';
        if ($find === '') {
            continue;
        }
        $valid[] = array(
            'find' => $find,
            'replace' => $replace,
        );
    }
    $stored = mdp_get_export_replace_rules_option();
    $stored[$form_id] = $valid;
    update_option('mdp_efs_export_replace_rules', $stored);
    wp_send_json_success();
}

function mdp_save_export_formulas_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    $payload_raw = isset($_POST['formulas_payload']) ? wp_unslash($_POST['formulas_payload']) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $payload = json_decode($payload_raw, true);
    $payload = is_array($payload) ? $payload : [];
    $valid = [];
    foreach ($payload as $entry) {
        if (!is_array($entry)) {
            continue;
        }
        $key = isset($entry['key']) ? sanitize_text_field($entry['key']) : '';
        $label = isset($entry['label']) ? sanitize_text_field($entry['label']) : '';
        $formula = isset($entry['formula']) ? sanitize_text_field($entry['formula']) : '';
        if ($key === '' || $formula === '') {
            continue;
        }
        if ($label === '') {
            $label = $key;
        }
        $valid[] = array(
            'key' => $key,
            'label' => $label,
            'formula' => $formula,
        );
    }
    $stored = mdp_get_export_formula_fields_option();
    $stored[$form_id] = $valid;
    update_option('mdp_efs_export_formula_fields', $stored);
    wp_send_json_success();
}

function mdp_get_export_formulas_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $formulas = mdp_get_export_formula_fields($form_id);
    wp_send_json_success(array('formulas' => $formulas));
}

function mdp_get_export_rules_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $rules = mdp_get_export_replace_rules($form_id);
    wp_send_json_success(array('rules' => $rules));
}

function mdp_get_export_preview_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    $preview_data = mdp_get_export_preview_data($form_id);
    wp_send_json_success($preview_data);
}

function mdp_save_export_last_form_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $form_id = isset($_POST['form_id']) ? sanitize_text_field(wp_unslash($_POST['form_id'])) : '';
    if ($form_id === '') {
        wp_send_json_error(array('message' => __('Kein Formular ausgewählt.', 'elementor-forms-statistics')));
    }
    mdp_set_last_export_form_id(get_current_user_id(), $form_id);
    wp_send_json_success();
}

function mdp_save_export_last_format_ajax() {
    if (!mdp_user_can_access_export_menu()) {
        wp_send_json_error(array('message' => __('Keine Berechtigung.', 'elementor-forms-statistics')));
    }
    check_ajax_referer('mdp_export_fields', 'nonce');
    $format = isset($_POST['format']) ? sanitize_text_field(wp_unslash($_POST['format'])) : '';
    mdp_set_last_export_format(get_current_user_id(), $format);
    wp_send_json_success();
}


function mdp_find_latest_submission_id($element_id) {
    global $wpdb;
    $element_id = sanitize_text_field($element_id);
    if ($element_id === '') {
        return 0;
    }
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $since = gmdate('Y-m-d H:i:s', current_time('timestamp', true) - 600);
    $user_ip = isset($_SERVER['REMOTE_ADDR']) ? sanitize_text_field(wp_unslash($_SERVER['REMOTE_ADDR'])) : '';
    $submission_id = (int) $wpdb->get_var($wpdb->prepare(
        "SELECT ID FROM {$submissions_table}
         WHERE element_id = %s
           AND created_at_gmt >= %s
           AND (%s = '' OR user_ip = %s)
         ORDER BY ID DESC
         LIMIT 1",
        $element_id,
        $since,
        $user_ip,
        $user_ip
    ));
    if ($submission_id > 0) {
        return $submission_id;
    }
    return (int) $wpdb->get_var($wpdb->prepare(
        "SELECT ID FROM {$submissions_table}
         WHERE element_id = %s
         ORDER BY ID DESC
         LIMIT 1",
        $element_id
    ));
}

function mdp_get_export_preview_data($form_id) {
    $form_id = sanitize_text_field($form_id);
    if ($form_id === '') {
        return array(
            'headers' => [],
            'rows' => [],
        );
    }
    $fields = array_values(array_filter(mdp_get_export_fields_for_form($form_id), function($field) {
        return !empty($field['include']);
    }));
    $formula_fields = mdp_get_export_formula_fields($form_id);
    $formula_map = [];
    foreach ($formula_fields as $formula_entry) {
        if (!empty($formula_entry['key'])) {
            $formula_map[$formula_entry['key']] = $formula_entry['formula'];
        }
    }
    $headers = [];
    foreach ($fields as $field) {
        $headers[] = isset($field['label']) && $field['label'] !== '' ? $field['label'] : $field['key'];
    }

    if (empty($fields)) {
        return array(
            'headers' => [],
            'rows' => [],
        );
    }

    global $wpdb;
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submission_rows = $wpdb->get_results($wpdb->prepare(
        "SELECT ID, created_at_gmt
         FROM {$submissions_table}
         WHERE element_id = %s
           AND status NOT LIKE %s
         ORDER BY created_at_gmt DESC",
        $form_id,
        '%trash%'
    ), ARRAY_A);
    if (empty($submission_rows)) {
        return array(
            'headers' => $headers,
            'rows' => [],
        );
    }

    $submission_ids = array_map('intval', wp_list_pluck($submission_rows, 'ID'));
    $created_at_map = [];
    foreach ($submission_rows as $submission_row) {
        $submission_id = isset($submission_row['ID']) ? (int) $submission_row['ID'] : 0;
        if ($submission_id <= 0) {
            continue;
        }
        $created_at_map[$submission_id] = mdp_format_submission_created_at($submission_row['created_at_gmt'] ?? '');
    }
    $values_by_submission = [];
    $chunk_size = 500;
    $field_keys = wp_list_pluck($fields, 'key');
    $allowed_keys = array_unique(array_merge($field_keys, (array) mdp_get_form_field_keys($form_id)));
    foreach (array_chunk($submission_ids, $chunk_size) as $chunk) {
        $placeholders = implode(',', array_fill(0, count($chunk), '%d'));
        $rows = $wpdb->get_results(
            $wpdb->prepare(
                "SELECT submission_id, `key`, `value`
                 FROM {$values_table}
                 WHERE submission_id IN ({$placeholders})",
                $chunk
            ),
            ARRAY_A
        );
        foreach ($rows as $row) {
            $submission_id = isset($row['submission_id']) ? (int) $row['submission_id'] : 0;
            $key = isset($row['key']) ? sanitize_text_field($row['key']) : '';
            if ($submission_id <= 0 || $key === '' || !in_array($key, $allowed_keys, true)) {
                continue;
            }
            $value = mdp_normalize_export_value($row['value']);
            if (!isset($values_by_submission[$submission_id])) {
                $values_by_submission[$submission_id] = [];
            }
            if (!isset($values_by_submission[$submission_id][$key])) {
                $values_by_submission[$submission_id][$key] = [];
            }
            if ($value !== '') {
                $values_by_submission[$submission_id][$key][] = $value;
            }
        }
    }

    $rows = [];
    foreach ($submission_ids as $submission_id) {
        if (!isset($values_by_submission[$submission_id])) {
            $values_by_submission[$submission_id] = [];
        }
        if (!isset($values_by_submission[$submission_id]['created_at']) && isset($created_at_map[$submission_id])) {
            $values_by_submission[$submission_id]['created_at'] = array($created_at_map[$submission_id]);
        }
        $value_map = mdp_get_submission_value_map($values_by_submission, $submission_id, $form_id);
        $row = [];
        foreach ($fields as $field) {
            $key = $field['key'];
            if (!empty($field['formula']) && isset($formula_map[$key])) {
                $row[] = mdp_apply_export_replace_rules($form_id, mdp_evaluate_export_formula($formula_map[$key], $value_map));
                continue;
            }
            $value = isset($value_map[$key]) ? $value_map[$key] : '';
            if (!empty($field['date_format'])) {
                $value = mdp_apply_export_date_format($value, $field['date_format']);
            }
            $row[] = mdp_apply_export_replace_rules($form_id, $value);
        }
        $rows[] = $row;
    }

    return array(
        'headers' => $headers,
        'rows' => $rows,
    );
}


// JS einfügen
function mdp_enqueue_chartjs($hook_suffix) {
    if ($hook_suffix !== 'toplevel_page_statistiken') {
        return;
    }
    wp_enqueue_script(
        'mdp-chartjs',
        plugins_url('/assets/js/chart.js', __FILE__),
        array(),
        '3.0.0',
        false
    );
}
add_action('admin_enqueue_scripts', 'mdp_enqueue_chartjs');

function mdp_enqueue_export_assets($hook_suffix) {
    if ($hook_suffix !== 'statistiken_page_statistiken-export') {
        return;
    }
    wp_enqueue_script('jquery-ui-sortable');
}
add_action('admin_enqueue_scripts', 'mdp_enqueue_export_assets');

function mdp_get_plugin_version_string() {
    static $cached = null;
    if ($cached !== null) {
        return $cached;
    }
    if (!function_exists('get_plugin_data')) {
        require_once ABSPATH . 'wp-admin/includes/plugin.php';
    }
    $plugin_data = get_plugin_data(__FILE__, false, false);
    $cached = isset($plugin_data['Version']) ? $plugin_data['Version'] : '';
    return $cached;
}

function custom_menu_callback($return_output = false, $render_options = array()) {
    if (!mdp_user_can_access_menu('statistiken')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }
    $render_defaults = array(
        'export' => false,
        'inline_styles' => false,
        'include_html_document' => false,
        'include_chartjs' => false,
        'show_inline_button' => true,
        'render_chart_as_image' => false,
    );
    $render_options = wp_parse_args($render_options, $render_defaults);
    if (!$return_output && is_admin()) {
        $render_options['show_inline_button'] = false;
    }

    $forms_indexed = mdp_get_forms_indexed();
    $display_mode = mdp_get_form_display_mode();
    $selected_form_ids = mdp_get_selected_form_ids($display_mode);
    $archive_ready = mdp_should_use_archive();
    $data_source = $archive_ready ? 'archive' : 'live';

    $dataset = mdp_collect_stats_dataset($data_source, array(
        'selected_form_ids' => $selected_form_ids,
        'forms_indexed' => $forms_indexed,
        'display_mode' => $display_mode,
    ));

    if ($archive_ready && empty($dataset['years_desc'])) {
        $archive_ready = false;
        $data_source = 'live';
        $dataset = mdp_collect_stats_dataset($data_source, array(
            'selected_form_ids' => $selected_form_ids,
            'forms_indexed' => $forms_indexed,
            'display_mode' => $display_mode,
        ));
    }

    $forms = $dataset['forms'];
    $form_order = $dataset['form_order'];
    $years_desc = $dataset['years_desc'];
    $years_asc = $dataset['years_asc'];
    $year_totals = $dataset['year_totals'];
    $current_year = $dataset['current_year'];
    $current_month = $dataset['current_month'];

    $month_template = mdp_get_empty_month_template();
    $month_labels = array(
        1 => __('Januar', 'elementor-forms-statistics'),
        2 => __('Februar', 'elementor-forms-statistics'),
        3 => __('März', 'elementor-forms-statistics'),
        4 => __('April', 'elementor-forms-statistics'),
        5 => __('Mai', 'elementor-forms-statistics'),
        6 => __('Juni', 'elementor-forms-statistics'),
        7 => __('Juli', 'elementor-forms-statistics'),
        8 => __('August', 'elementor-forms-statistics'),
        9 => __('September', 'elementor-forms-statistics'),
        10 => __('Oktober', 'elementor-forms-statistics'),
        11 => __('November', 'elementor-forms-statistics'),
        12 => __('Dezember', 'elementor-forms-statistics'),
    );

    $color_slots = mdp_get_curve_color_slots();
    $slot_count = count($color_slots);

    $reference_year = !empty($years_asc) ? (int) max($years_asc) : (int) date('Y');
    $chart_datasets = array();
    foreach ($years_asc as $index => $year) {
        $year_diff = max(0, $reference_year - (int) $year);
        $slot_index = min($year_diff, $slot_count - 1);
        $slot = $color_slots[$slot_index];
        $border_color = mdp_hex_to_rgba($slot['color'], $slot['alpha']);
        $fill_color = mdp_hex_to_rgba($slot['color'], max(0.05, $slot['alpha'] * 0.45));
        $data_points = array();
        foreach ($month_labels as $month_index => $label) {
            $value = isset($year_totals[$year][$month_index]) ? (int) $year_totals[$year][$month_index] : 0;
            $data_points[] = $value;
        }
        $chart_datasets[] = array(
            'label' => (string) $year,
            'data' => $data_points,
            'borderColor' => $border_color,
            'pointBackgroundColor' => $border_color,
            'pointBorderColor' => $border_color,
            'borderWidth' => 3,
            'backgroundColor' => $fill_color,
            'fill' => true,
        );
    }
    $datasets_json = wp_json_encode($chart_datasets);
    $has_chart = !empty($chart_datasets);

    $has_data = !empty($form_order) && !empty($years_desc);

    ob_start();

    if ($render_options['include_html_document']) {
        echo '<!DOCTYPE html><html><head>';
        echo '<meta charset="' . esc_attr(get_bloginfo('charset')) . '">';
        echo '<title>' . esc_html__('Anfragen Statistik', 'elementor-forms-statistics') . '</title>';
        echo '<meta name="robots" content="noindex,nofollow,noarchive">';
        if ($render_options['inline_styles']) {
            echo '<style>' . mdp_get_inline_export_styles() . '</style>';
        }
        echo '</head><body class="mdp-export-body">';
    } else {
        if ($render_options['inline_styles']) {
            echo '<style>' . mdp_get_inline_export_styles() . '</style>';
        }
    }

    echo '<div class="wrap mdp-stats-root">';
    $site_url = home_url();
    $site_url_display = preg_replace('#^https?://#', '', $site_url);
    echo '<h1>' . esc_html(sprintf(__('Statistiken %s', 'elementor-forms-statistics'), $site_url_display)) . '</h1>';
    if (!$render_options['export'] && $render_options['show_inline_button']) {
        $inline_url = wp_nonce_url(
            add_query_arg(
                array(
                    'action' => 'mdp_export_stats_html',
                    'inline' => 1,
                ),
                admin_url('admin-post.php')
            ),
            'mdp_export_html_inline',
            'mdp_inline_nonce'
        );
        echo '<div class="mdp-inline-view">';
        echo '<span class="mdp-inline-text">' . esc_html__('Für Chartansicht im Browser ansehen', 'elementor-forms-statistics') . '</span>';
        echo '<a class="button button-primary mdp-inline-button" href="' . esc_url($inline_url) . '" target="_blank" rel="noopener noreferrer">';
        echo esc_html__('Im Browser anzeigen', 'elementor-forms-statistics');
        echo '</a></div>';
    }

    if ($has_chart) {
        if ($render_options['render_chart_as_image']) {
            $chart_image_data = mdp_generate_chart_image_base64($chart_datasets, $current_year, $current_month);
            if ($chart_image_data) {
                echo '<div class="mdp-chart mdp-chart-image"><img src="' . esc_attr($chart_image_data) . '" alt="' . esc_attr__('Diagramm der Formularanfragen', 'elementor-forms-statistics') . '"></div>';
            } else {
                echo '<p>' . esc_html__('Das Diagramm konnte nicht als Bild gerendert werden.', 'elementor-forms-statistics') . '</p>';
            }
        } else {
            echo '<div class="mdp-chart"><canvas id="myChart"></canvas></div>';
        }
    } else {
        echo '<p>' . esc_html__('Keine Daten vorhanden.', 'elementor-forms-statistics') . '</p>';
    }

    $plugin_version = mdp_get_plugin_version_string();
    $footer_text = sprintf(__('%s · 2025 · v%s', 'elementor-forms-statistics'), __('Elementor Forms Statistics', 'elementor-forms-statistics'), $plugin_version);
    $footer_link = '<a href="https://www.medienproduktion.biz/elementor-forms-statistics/" target="_blank" rel="noopener noreferrer">' . esc_html($footer_text) . '</a>';
    echo '<div class="mdp-chart-footer">' . $footer_link . '</div>';

    if ($has_data) {
        foreach ($years_desc as $year) {
            echo '<h2 class="mdp-year-heading">' . esc_html($year) . '</h2>';
            echo '<div class="mdp-table-wrapper">';
            echo '<table class="mdp-table"><thead><tr>';
            echo '<th>' . esc_html__('Formular', 'elementor-forms-statistics') . '</th>';
            foreach ($month_labels as $month_label) {
                echo '<th>' . esc_html($month_label) . '</th>';
            }
            echo '<th>' . esc_html__('Summe', 'elementor-forms-statistics') . '</th>';
            echo '</tr></thead><tbody>';

            foreach ($form_order as $form_id) {
                if (!isset($forms[$form_id])) {
                    continue;
                }
                $form_data = $forms[$form_id];
                $current_months = isset($form_data['yearly'][$year]) ? $form_data['yearly'][$year] : $month_template;
                $previous_year_months = isset($form_data['yearly'][$year - 1]) ? $form_data['yearly'][$year - 1] : $month_template;
                $has_previous_year_data = isset($form_data['yearly'][$year - 1]);
                $annual_total = 0;
                $previous_annual_total = 0;
                foreach ($current_months as $month_value) {
                    $annual_total += (int) $month_value;
                }
                foreach ($previous_year_months as $prev_value) {
                    $previous_annual_total += (int) $prev_value;
                }

                $row_title = $form_data['title'] !== '' ? $form_data['title'] : __('Unnamed Form', 'elementor-forms-statistics');
                echo '<tr data-form-id="' . esc_attr($form_id) . '">';
                echo '<td>' . esc_html($row_title) . '</td>';

                foreach ($month_labels as $month_index => $month_label) {
                    $value = isset($current_months[$month_index]) ? (int) $current_months[$month_index] : 0;
                    $previous_value = isset($previous_year_months[$month_index]) ? (int) $previous_year_months[$month_index] : 0;
                    $is_past_period = ($year < $current_year) || ($year === $current_year && $month_index < $current_month);
                    $cell_style = $is_past_period ? mdp_get_comparison_background($value, $previous_value, $has_previous_year_data) : '';
                    $emphasize_change = $is_past_period ? mdp_should_emphasize_change($value, $previous_value, $has_previous_year_data) : false;

                    if ($emphasize_change) {
                        $content = '<strong>' . ($value === 0 ? '0' : esc_html($value)) . '</strong>';
                    } elseif ($value === 0) {
                        $content = '&nbsp;';
                    } else {
                        $content = esc_html($value);
                    }
                    echo '<td' . $cell_style . '>' . $content . '</td>';
                }

                $annual_is_past = ($year < $current_year);
                $annual_cell_style = $annual_is_past ? mdp_get_comparison_background($annual_total, $previous_annual_total, $has_previous_year_data) : '';
                $annual_emphasis = $annual_is_past ? mdp_should_emphasize_change($annual_total, $previous_annual_total, $has_previous_year_data) : false;
                $annual_content = $annual_emphasis ? '<strong>' . esc_html($annual_total) . '</strong>' : esc_html($annual_total);
                echo '<td' . $annual_cell_style . '>' . $annual_content . '</td>';
                echo '</tr>';
            }

            $current_year_totals = isset($year_totals[$year]) ? $year_totals[$year] : $month_template;
            $previous_year_totals = isset($year_totals[$year - 1]) ? $year_totals[$year - 1] : $month_template;
            $has_previous_year_totals = isset($year_totals[$year - 1]);
            $grand_total = array_sum($current_year_totals);
            $previous_grand_total = array_sum($previous_year_totals);

            echo '<tr>';
            echo '<td><strong>' . esc_html__('Gesamt', 'elementor-forms-statistics') . '</strong></td>';
            foreach ($month_labels as $month_index => $month_label) {
                $value = isset($current_year_totals[$month_index]) ? (int) $current_year_totals[$month_index] : 0;
                $previous_value = isset($previous_year_totals[$month_index]) ? (int) $previous_year_totals[$month_index] : 0;
                $is_past_period = ($year < $current_year) || ($year === $current_year && $month_index < $current_month);
                $cell_style = $is_past_period ? mdp_get_comparison_background($value, $previous_value, $has_previous_year_totals) : '';
                $emphasize_change = $is_past_period ? mdp_should_emphasize_change($value, $previous_value, $has_previous_year_totals) : false;
                $content_value = $emphasize_change ? '<strong>' . esc_html($value) . '</strong>' : '<strong>' . esc_html($value) . '</strong>';
                echo '<td' . $cell_style . '>' . $content_value . '</td>';
            }
            $grand_is_past = ($year < $current_year);
            $grand_cell_style = $grand_is_past ? mdp_get_comparison_background($grand_total, $previous_grand_total, $has_previous_year_totals) : '';
            $grand_emphasis = $grand_is_past ? mdp_should_emphasize_change($grand_total, $previous_grand_total, $has_previous_year_totals) : false;
            $grand_content = $grand_emphasis ? '<strong>' . esc_html($grand_total) . '</strong>' : '<strong>' . esc_html($grand_total) . '</strong>';
            echo '<td' . $grand_cell_style . '>' . $grand_content . '</td>';
            echo '</tr>';

            echo '</tbody></table>';
            echo '</div>';
        }
    }

    echo '</div>';

    if ($has_chart && $render_options['include_chartjs'] && !$render_options['render_chart_as_image']) {
        $chart_js_path = plugin_dir_path(__FILE__) . 'assets/js/chart.js';
        if (file_exists($chart_js_path)) {
            $chart_js_content = file_get_contents($chart_js_path);
            if ($chart_js_content !== false) {
                echo '<script>' . $chart_js_content . '</script>';
            }
        }
    }

    if ($has_chart && !$render_options['render_chart_as_image']) {
        ?>
        <script>
        (function() {
            var canvas = document.getElementById('myChart');
            if (!canvas) {
                return;
            }
            var ctx = canvas.getContext('2d');
            var chartData = <?php echo $datasets_json; ?>;
            var currentYear = <?php echo (int) $current_year; ?>;
            var currentMonthIndex = <?php echo (int) $current_month - 1; ?>;

            chartData.forEach(function(dataset) {
                if (parseInt(dataset.label, 10) === currentYear) {
                    dataset.segment = dataset.segment || {};
                    dataset.segment.borderDash = function(ctx) {
                        if (ctx.p0DataIndex === currentMonthIndex || ctx.p1DataIndex === currentMonthIndex) {
                            return [6, 4];
                        }
                        return undefined;
                    };
                }
            });

            var labels = [];
            for (var month = 1; month <= 12; month++) {
                labels.push(month);
            }

            var data = {
                labels: labels,
                datasets: chartData
            };

            var options = {
                responsive: true,
                maintainAspectRatio: false,
                tooltips: {
                    callbacks: {
                        title: function() {
                            return '<?php echo esc_js(__('Anzahl Anfragen', 'elementor-forms-statistics')); ?>';
                        },
                        label: function(tooltipItem) {
                            return tooltipItem.yLabel + ' ' + '<?php echo esc_js(__('Anfragen', 'elementor-forms-statistics')); ?>';
                        }
                    }
                },
                scales: {
                    y: {
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            };
            options.onHover = function(event, elements) {
                var canvas = this && this.canvas ? this.canvas : (event && event.target ? event.target : null);
                if (canvas) {
                    canvas.style.cursor = elements && elements.length ? 'pointer' : 'default';
                }
            };

            function hexToRgba(hex, alpha) {
                var r = parseInt(hex.substr(1, 2), 16);
                var g = parseInt(hex.substr(3, 2), 16);
                var b = parseInt(hex.substr(5, 2), 16);
                return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha + ')';
            }

            for (var i = 0; i < chartData.length; i++) {
                chartData[i].fill = true;
                chartData[i].tension = 0.35;
                var lineColor = chartData[i].borderColor || '#4b5563';
                if (lineColor.indexOf('rgba') === 0) {
                    var rgbaParts = lineColor.match(/^rgba\((\d+),\s*(\d+),\s*(\d+),\s*(\d+\.?\d*)\)$/);
                    if (rgbaParts) {
                        chartData[i].backgroundColor = 'rgba(' + rgbaParts[1] + ',' + rgbaParts[2] + ',' + rgbaParts[3] + ',0.1)';
                    }
                } else if (lineColor.indexOf('#') === 0) {
                    chartData[i].backgroundColor = hexToRgba(lineColor, 0.1);
                }
            }

            var chartInstance = new Chart(ctx, {
                type: 'line',
                data: data,
                options: options
            });
        })();
        </script>
        <?php
    }

    if ($render_options['include_html_document']) {
        echo '</body></html>';
    }

    $output = ob_get_clean();
    if ($return_output) {
        return $output;
    }
    echo $output;
}

function mdp_hex_to_rgb_array($hex_color) {
    $hex_color = trim($hex_color);
    if ($hex_color === '') {
        return array(0, 0, 0);
    }
    if ($hex_color[0] === '#') {
        $hex_color = substr($hex_color, 1);
    }
    if (strlen($hex_color) === 3) {
        $hex_color = $hex_color[0] . $hex_color[0] . $hex_color[1] . $hex_color[1] . $hex_color[2] . $hex_color[2];
    }
    $int = hexdec($hex_color);
    return array(
        ($int >> 16) & 255,
        ($int >> 8) & 255,
        $int & 255,
    );
}

function mdp_catmull_rom_point($p0, $p1, $p2, $p3, $t, $tension) {
    $t2 = $t * $t;
    $t3 = $t2 * $t;
    $v0x = ($p2['x'] - $p0['x']) * $tension;
    $v0y = ($p2['y'] - $p0['y']) * $tension;
    $v1x = ($p3['x'] - $p1['x']) * $tension;
    $v1y = ($p3['y'] - $p1['y']) * $tension;

    $x = (2 * $p1['x'] - 2 * $p2['x'] + $v0x + $v1x) * $t3
        + (-3 * $p1['x'] + 3 * $p2['x'] - 2 * $v0x - $v1x) * $t2
        + $v0x * $t + $p1['x'];

    $y = (2 * $p1['y'] - 2 * $p2['y'] + $v0y + $v1y) * $t3
        + (-3 * $p1['y'] + 3 * $p2['y'] - 2 * $v0y - $v1y) * $t2
        + $v0y * $t + $p1['y'];

    return array('x' => $x, 'y' => $y);
}

function mdp_generate_spline_points($points, $segments = 12, $tension = 0.35) {
    $count = count($points);
    if ($count < 3) {
        return $points;
    }
    $spline = array();
    for ($i = 0; $i < $count - 1; $i++) {
        $p0 = $points[$i === 0 ? $i : $i - 1];
        $p1 = $points[$i];
        $p2 = $points[$i + 1];
        $p3 = $points[$i + 1 >= $count - 1 ? $count - 1 : $i + 2];

        for ($j = 0; $j < $segments; $j++) {
            $t = $j / $segments;
            $spline[] = mdp_catmull_rom_point($p0, $p1, $p2, $p3, $t, $tension);
        }
    }
    $spline[] = end($points);
    return $spline;
}

function mdp_generate_chart_image_base64($chart_datasets, $current_year, $current_month) {
    if (empty($chart_datasets) || !function_exists('imagecreatetruecolor') || !function_exists('imagepng')) {
        return '';
    }

    $width = 1090;
    $height = 420;
    $padding = 18;
    $chart_left = 70;
    $chart_top = 40;
    $chart_right = $width - 30;
    $chart_bottom = $height - 60;
    $chart_width = $chart_right - $chart_left;
    $chart_height = $chart_bottom - $chart_top;
    $months = 12;
    $month_spacing = $chart_width / ($months - 1);

    $max_value = 0;
    foreach ($chart_datasets as $dataset) {
        if (empty($dataset['data']) || !is_array($dataset['data'])) {
            continue;
        }
        foreach ($dataset['data'] as $value) {
            if ($value > $max_value) {
                $max_value = $value;
            }
        }
    }
    if ($max_value < 1) {
        $max_value = 1;
    }

    $image = imagecreatetruecolor($width, $height);
    imagealphablending($image, true);
    imagesavealpha($image, true);

    $background = imagecolorallocate($image, 231, 234, 239);
    $card_fill = imagecolorallocate($image, 248, 249, 252);
    $card_border = imagecolorallocate($image, 220, 224, 232);
    $grid_color = imagecolorallocate($image, 214, 218, 226);
    $axis_color = imagecolorallocate($image, 170, 178, 192);
    $text_color = imagecolorallocate($image, 29, 35, 39);
    $legend_text_color = imagecolorallocate($image, 50, 57, 65);
    imagefilledrectangle($image, 0, 0, $width - 1, $height - 1, $background);
    imagefilledrectangle($image, $padding, $padding, $width - $padding, $height - $padding, $card_fill);
    imagerectangle($image, $padding, $padding, $width - $padding, $height - $padding, $card_border);
    if (function_exists('imageantialias')) {
        imageantialias($image, true);
    }

    $title = __('Anfragen pro Monat und Jahr', 'elementor-forms-statistics');
    imagestring($image, 5, $chart_left, 12, $title, $text_color);

    $grid_steps = 6;
    $step_value = $max_value / $grid_steps;
    for ($i = 0; $i <= $grid_steps; $i++) {
        $y = $chart_bottom - ($chart_height / $grid_steps) * $i;
        imageline($image, $chart_left, (int) $y, $chart_right, (int) $y, $grid_color);
        $label_value = round($step_value * $i);
        imagestring($image, 2, 20, (int) $y - 6, (string) $label_value, $text_color);
    }

    for ($i = 0; $i < $months; $i++) {
        $x = $chart_left + $month_spacing * $i;
        imageline($image, (int) $x, $chart_top, (int) $x, $chart_bottom, $grid_color);
        imagestring($image, 2, (int) ($x - 5), $chart_bottom + 10, (string) ($i + 1), $text_color);
    }

    imageline($image, $chart_left, $chart_top, $chart_left, $chart_bottom, $axis_color);
    imageline($image, $chart_left, $chart_bottom, $chart_right, $chart_bottom, $axis_color);

    $legend_x = $chart_left;
    $legend_y = 30;
    $legend_gap = 120;

    foreach ($chart_datasets as $dataset) {
        if (empty($dataset['data']) || !is_array($dataset['data'])) {
            continue;
        }
        $rgb = mdp_hex_to_rgb_array(isset($dataset['borderColor']) ? $dataset['borderColor'] : '#2d3748');
        $line_color = imagecolorallocate($image, $rgb[0], $rgb[1], $rgb[2]);
        $fill_color = imagecolorallocatealpha($image, $rgb[0], $rgb[1], $rgb[2], 110);
        imagesetthickness($image, 3);

        imagefilledrectangle($image, $legend_x, $legend_y - 12, $legend_x + 14, $legend_y + 2, $line_color);
        $label = isset($dataset['label']) ? $dataset['label'] : '';
        imagestring($image, 4, $legend_x + 20, $legend_y - 12, (string) $label, $legend_text_color);
        $legend_x += $legend_gap;
        if ($legend_x + $legend_gap > $chart_right) {
            $legend_x = $chart_left;
            $legend_y += 20;
        }

        $points = array();
        foreach ($dataset['data'] as $index => $value) {
            $x = $chart_left + $month_spacing * $index;
            $value = max(0, (float) $value);
            $ratio = $value / $max_value;
            $y = $chart_bottom - ($ratio * $chart_height);
            $points[] = array(
                'x' => $x,
                'y' => $y,
                'month_index' => $index,
            );
        }

        if (count($points) < 2) {
            continue;
        }

        $spline_points = mdp_generate_spline_points($points, 10, 0.35);
        $polygon = array();
        foreach ($spline_points as $pt) {
            $polygon[] = (int) round($pt['x']);
            $polygon[] = (int) round($pt['y']);
        }
        $last_point = end($points);
        $first_point = reset($points);
        $polygon[] = (int) round($last_point['x']);
        $polygon[] = (int) round($chart_bottom);
        $polygon[] = (int) round($first_point['x']);
        $polygon[] = (int) round($chart_bottom);
        imagefilledpolygon($image, $polygon, count($polygon) / 2, $fill_color);

        $prev = null;
        $dash_start = $chart_left + $month_spacing * max($current_month - 2, 0);
        $dash_end = $chart_left + $month_spacing * min($current_month, $months - 1);
        foreach ($spline_points as $pt) {
            if ($prev) {
                $x1 = (int) round($prev['x']);
                $y1 = (int) round($prev['y']);
                $x2 = (int) round($pt['x']);
                $y2 = (int) round($pt['y']);
                $use_dash = isset($dataset['label']) && (string) $dataset['label'] === (string) $current_year
                    && (($x1 >= $dash_start && $x1 <= $dash_end) || ($x2 >= $dash_start && $x2 <= $dash_end));
                if ($use_dash) {
                    imagedashedline($image, $x1, $y1, $x2, $y2, $line_color);
                } else {
                    imageline($image, $x1, $y1, $x2, $y2, $line_color);
                }
            }
            $prev = $pt;
        }
    }

    ob_start();
    imagepng($image);
    $png_data = ob_get_clean();
    imagedestroy($image);
    if ($png_data === false || $png_data === '') {
        return '';
    }
    return 'data:image/png;base64,' . base64_encode($png_data);
}

// Fetch all unique forms using the stable element_id as identifier
function custom_get_all_forms() {
    global $wpdb;

    if (mdp_archive_table_exists() && mdp_archive_has_data()) {
        $archive_table = mdp_get_archive_table_name();
        $rows = $wpdb->get_results("
            SELECT form_id AS element_id,
                   MAX(NULLIF(form_name, '')) AS form_title
            FROM {$archive_table}
            GROUP BY form_id
            ORDER BY form_title
        ");
        if (!empty($rows)) {
            return $rows;
        }
    }

    return $wpdb->get_results(
        "SELECT element_id,
                MAX(NULLIF(form_name, '')) AS form_title
        FROM " . $wpdb->prefix . "e_submissions
        WHERE element_id IS NOT NULL AND element_id != ''
        GROUP BY element_id
        ORDER BY form_title"
    );
}

function mdp_get_form_display_mode() {
    $mode = get_option('mdp_efs_form_display_mode', 'element_id');
    $allowed = array('element_id', 'form_name', 'form_name_distinct', 'referer_title', 'referer_title_distinct');
    return in_array($mode, $allowed, true) ? $mode : 'element_id';
}

function mdp_get_selected_form_ids($mode = null) {
    $stored = get_option('mdp_efs_selected_form_ids', []);
    $mode = $mode ? $mode : mdp_get_form_display_mode();
    if (!is_array($stored)) {
        return [];
    }
    if (isset($stored['element_id']) || isset($stored['form_name']) || isset($stored['form_name_distinct'])) {
        $list = isset($stored[$mode]) && is_array($stored[$mode]) ? $stored[$mode] : [];
        return array_map('sanitize_text_field', $list);
    }
    $flat = array_map('sanitize_text_field', $stored);
    return $mode === 'element_id' ? $flat : [];
}

function mdp_get_forms_indexed() {
    $forms = custom_get_all_forms();
    $indexed = [];
    foreach ($forms as $form_entry) {
        if (!empty($form_entry->element_id)) {
            $form_entry->form_title = isset($form_entry->form_title) ? sanitize_text_field($form_entry->form_title) : '';
            $indexed[$form_entry->element_id] = $form_entry;
        }
    }

    $archived_forms = mdp_get_archived_form_titles();
    if (!empty($archived_forms)) {
        foreach ($archived_forms as $archived) {
            if (empty($archived->form_id)) {
                continue;
            }
            if (!isset($indexed[$archived->form_id]) || empty($indexed[$archived->form_id]->referer_title)) {
                $indexed[$archived->form_id] = (object) array(
                    'element_id' => $archived->form_id,
                    'form_title' => sanitize_text_field($archived->form_title),
                );
            }
        }
    }

    $snapshot_titles = mdp_get_elementor_form_snapshot_titles();
    if (!empty($snapshot_titles)) {
        foreach ($snapshot_titles as $form_id => $form_title) {
            if (isset($indexed[$form_id])) {
                $indexed[$form_id]->form_title = $form_title;
            } else {
                $indexed[$form_id] = (object) array(
                    'element_id' => $form_id,
                    'form_title' => $form_title,
                );
            }
        }
    }

    return $indexed;
}

function mdp_resolve_form_title($form_id, $form_entry = null) {
    if ($form_entry && !empty($form_entry->form_title)) {
        return $form_entry->form_title;
    }

    static $resolved_titles = [];
    if (isset($resolved_titles[$form_id])) {
        return $resolved_titles[$form_id];
    }

    global $wpdb;
    $row = $wpdb->get_row($wpdb->prepare(
        "SELECT form_name FROM " . $wpdb->prefix . "e_submissions
         WHERE element_id = %s
         ORDER BY created_at_gmt DESC LIMIT 1",
        $form_id
    ), ARRAY_A);

    $title = '';
    if ($row) {
        if (!empty($row['form_name'])) {
            $title = sanitize_text_field($row['form_name']);
        }
    }

    $resolved_titles[$form_id] = $title;
    return $resolved_titles[$form_id];
}

function mdp_get_empty_month_template() {
    return array(
        1 => 0,
        2 => 0,
        3 => 0,
        4 => 0,
        5 => 0,
        6 => 0,
        7 => 0,
        8 => 0,
        9 => 0,
        10 => 0,
        11 => 0,
        12 => 0,
    );
}

function mdp_format_form_title($form_id, $forms_indexed, $fallback = '') {
    if (isset($forms_indexed[$form_id]) && !empty($forms_indexed[$form_id]->referer_title)) {
        return $forms_indexed[$form_id]->referer_title;
    }
    if ($fallback !== '') {
        return $fallback;
    }
    $resolved = mdp_resolve_form_title($form_id);
    if ($resolved !== '') {
        return $resolved;
    }
    return __('Unbenannte Seite', 'elementor-forms-statistics');
}

function mdp_get_form_title_overrides() {
    $stored = get_option('mdp_efs_form_title_overrides', []);
    if (!is_array($stored)) {
        return [];
    }
    $overrides = [];
    foreach ($stored as $form_id => $title) {
        $form_id = sanitize_text_field($form_id);
        if ($form_id === '') {
            continue;
        }
        if (!is_string($title)) {
            continue;
        }
        $overrides[$form_id] = sanitize_text_field($title);
    }
    return $overrides;
}

function mdp_build_canonical_form_mapping($forms_indexed, $title_overrides = array(), $display_mode = 'element_id') {
    $canonical_forms = array();
    $element_to_canonical = array();
    $unnamed_label = __('Unnamed Form', 'elementor-forms-statistics');

    foreach ($forms_indexed as $form_id => $form_entry) {
        $resolved_title = '';
        if (isset($form_entry->form_title) && $form_entry->form_title !== '') {
            $resolved_title = $form_entry->form_title;
        } else {
            $resolved_title = mdp_resolve_form_title($form_id, $form_entry);
        }
        $display_title = $resolved_title !== '' ? $resolved_title : $unnamed_label;

        $canonical_forms[$form_id] = array(
            'title' => $display_title,
            'default_title' => $display_title,
            'member_ids' => array($form_id),
            'title_options' => array(array(
                'title' => $display_title,
                'ids' => array($form_id),
            )),
            'override_title' => isset($title_overrides[$form_id]) ? $title_overrides[$form_id] : '',
        );
        $element_to_canonical[$form_id] = $form_id;
    }

    foreach ($canonical_forms as $form_id => &$entry) {
        $override_value = isset($title_overrides[$form_id]) ? $title_overrides[$form_id] : '';
        $existing_titles = wp_list_pluck($entry['title_options'], 'title');
        if ($override_value !== '' && !in_array($override_value, $existing_titles, true)) {
            $entry['title_options'][] = array(
                'title' => $override_value,
                'ids' => array(),
            );
        }
        $selected_title = $override_value !== '' ? $override_value : $entry['default_title'];
        if ($selected_title === '' && !empty($entry['title_options'])) {
            $selected_title = $entry['title_options'][0]['title'];
        }
        $entry['title'] = $selected_title !== '' ? $selected_title : $unnamed_label;
        $entry['override_title'] = $override_value;
    }
    unset($entry);

    if ($display_mode === 'form_name_distinct') {
        $grouped_forms = array();
        $new_element_map = array();
        foreach ($canonical_forms as $element_id => $entry) {
            $group_title = $entry['title'] !== '' ? $entry['title'] : $unnamed_label;
            $normalized = trim($group_title);
            $lower = function_exists('mb_strtolower') ? mb_strtolower($normalized) : strtolower($normalized);
            $group_key = 'formname_' . md5($lower);
            if (!isset($grouped_forms[$group_key])) {
                $grouped_forms[$group_key] = array(
                    'title' => $group_title,
                    'default_title' => $group_title,
                    'member_ids' => array(),
                    'title_options' => array(),
                    'override_title' => '',
                );
            }
            $grouped_forms[$group_key]['member_ids'] = array_merge($grouped_forms[$group_key]['member_ids'], $entry['member_ids']);
            $grouped_forms[$group_key]['title_options'] = array_merge($grouped_forms[$group_key]['title_options'], $entry['title_options']);
            if (!empty($entry['override_title'])) {
                $grouped_forms[$group_key]['override_title'] = $entry['override_title'];
            }
            $new_element_map[$element_id] = $group_key;
        }
        foreach ($grouped_forms as $group_key => &$group_entry) {
            $group_entry['member_ids'] = array_values(array_unique($group_entry['member_ids']));
            $existing_titles = wp_list_pluck($group_entry['title_options'], 'title');
            if (!empty($existing_titles)) {
                $group_entry['title'] = reset($existing_titles);
                $group_entry['default_title'] = reset($existing_titles);
            }
        }
        unset($group_entry);
        $canonical_forms = $grouped_forms;
        $element_to_canonical = $new_element_map;
    }

    return array(
        'canonical_forms' => $canonical_forms,
        'element_to_canonical' => $element_to_canonical,
    );
}

function mdp_get_submission_counts_by_form() {
    global $wpdb;
    $archive_counts = mdp_get_submission_counts_from_archive();
    if (!empty($archive_counts)) {
        return $archive_counts;
    }

    $table = $wpdb->prefix . 'e_submissions';
    $rows = $wpdb->get_results("
        SELECT element_id, COUNT(*) AS total
        FROM {$table}
        WHERE element_id IS NOT NULL
          AND element_id != ''
          AND status NOT LIKE '%trash%'
        GROUP BY element_id
    ", ARRAY_A);
    $counts = [];
    foreach ($rows as $row) {
        if (empty($row['element_id'])) {
            continue;
        }
        $counts[$row['element_id']] = (int) $row['total'];
    }
    return $counts;
}

function mdp_get_submission_counts_from_archive() {
    if (!mdp_archive_table_exists()) {
        return [];
    }
    global $wpdb;
    $table = mdp_get_archive_table_name();
    $rows = $wpdb->get_results("
        SELECT form_id, SUM(total) AS total
        FROM {$table}
        GROUP BY form_id
    ", ARRAY_A);
    $counts = [];
    foreach ($rows as $row) {
        if (empty($row['form_id'])) {
            continue;
        }
        $counts[$row['form_id']] = (int) $row['total'];
    }
    return $counts;
}

function mdp_get_archive_form_rows($display_mode) {
    if (!mdp_archive_table_exists()) {
        return [];
    }
    global $wpdb;
    $table = mdp_get_archive_table_name();
    $name_expr = $display_mode === 'referer_title' || $display_mode === 'referer_title_distinct'
        ? "NULLIF(referer_title, '')"
        : "NULLIF(form_name, '')";
    $rows = [];
    if ($display_mode === 'form_name' || $display_mode === 'referer_title') {
        $rows = $wpdb->get_results("
            SELECT form_id,
                   {$name_expr} AS form_name,
                   SUM(total) AS total
            FROM {$table}
            GROUP BY form_id, form_name
        ", ARRAY_A);
    } elseif ($display_mode === 'form_name_distinct' || $display_mode === 'referer_title_distinct') {
        $rows = $wpdb->get_results("
            SELECT MD5(LOWER(TRIM({$name_expr}))) AS form_key,
                   {$name_expr} AS form_name,
                   GROUP_CONCAT(DISTINCT form_id ORDER BY form_id SEPARATOR ',') AS ids,
                   SUM(total) AS total
            FROM {$table}
            GROUP BY form_key, form_name
        ", ARRAY_A);
    } else {
        $rows = $wpdb->get_results("
            SELECT form_id,
                   MAX({$name_expr}) AS form_name,
                   SUM(total) AS total
            FROM {$table}
            GROUP BY form_id
        ", ARRAY_A);
    }

    $results = array();
    foreach ($rows as $row) {
        $raw_title = isset($row['form_name']) ? trim((string) $row['form_name']) : '';
        $display_title = $raw_title !== '' ? $raw_title : __('Unbenannte Seite', 'elementor-forms-statistics');
        if ($display_mode === 'form_name_distinct') {
            $key_source = $raw_title !== '' ? $raw_title : $display_title;
            $lower = function_exists('mb_strtolower') ? mb_strtolower($key_source) : strtolower($key_source);
            $key = !empty($row['form_key']) ? $row['form_key'] : md5($lower);
            $ids = !empty($row['ids']) ? array_filter(array_map('trim', explode(',', $row['ids']))) : array();
        } elseif ($display_mode === 'form_name' || $display_mode === 'referer_title') {
            $form_id = isset($row['form_id']) ? (string) $row['form_id'] : '';
            $key = $form_id . '||' . ($raw_title !== '' ? $raw_title : $display_title);
            $ids = $form_id !== '' ? array($form_id) : array();
        } else {
            $form_id = isset($row['form_id']) ? (string) $row['form_id'] : '';
            $key = $form_id;
            $ids = $form_id !== '' ? array($form_id) : array();
        }
        if ($key === '') {
            continue;
        }
        $results[$key] = array(
            'title' => sanitize_text_field($display_title),
            'ids' => array_map('sanitize_text_field', $ids),
            'count' => isset($row['total']) ? (int) $row['total'] : 0,
        );
    }
    return $results;
}

function mdp_collect_stats_dataset($data_source = 'live', $args = array()) {
    global $wpdb;

    $defaults = array(
        'selected_form_ids' => array(),
        'forms_indexed' => array(),
    );
    $args = wp_parse_args($args, $defaults);
    $display_mode = isset($args['display_mode']) ? $args['display_mode'] : mdp_get_form_display_mode();
    $selected_ids = array();
    foreach ($args['selected_form_ids'] as $form_id) {
        $form_id = sanitize_text_field($form_id);
        if ($form_id !== '') {
            $selected_ids[] = $form_id;
        }
    }

    if ($display_mode !== 'element_id') {
        $month_template = mdp_get_empty_month_template();
        $selected_keys = array_values(array_unique(array_filter($selected_ids)));
        $rows = array();
        if ($data_source === 'archive' && mdp_archive_table_exists()) {
            $table = mdp_get_archive_table_name();
            $name_expr = $display_mode === 'referer_title' || $display_mode === 'referer_title_distinct'
                ? "NULLIF(referer_title, '')"
                : "NULLIF(form_name, '')";
            $where = "WHERE 1=1";
            if (!empty($selected_keys)) {
                $placeholders = implode(', ', array_fill(0, count($selected_keys), '%s'));
                if ($display_mode === 'form_name' || $display_mode === 'referer_title') {
                    $where .= $wpdb->prepare(" AND CONCAT(form_id,'||',{$name_expr}) IN ($placeholders)", $selected_keys);
                } else {
                    $where .= $wpdb->prepare(" AND MD5(LOWER(TRIM({$name_expr}))) IN ($placeholders)", $selected_keys);
                }
            }
            if ($display_mode === 'form_name' || $display_mode === 'referer_title') {
                $rows = $wpdb->get_results("
                    SELECT CONCAT(form_id,'||',{$name_expr}) AS form_key,
                           {$name_expr} AS form_name,
                           year AS year_value,
                           month AS month_value,
                           SUM(total) AS total
                    FROM {$table}
                    {$where}
                    GROUP BY form_id, form_name, year_value, month_value
                    ORDER BY year_value, month_value
                ");
            } else {
                $rows = $wpdb->get_results("
                    SELECT MD5(LOWER(TRIM({$name_expr}))) AS form_key,
                           MIN({$name_expr}) AS form_name,
                           year AS year_value,
                           month AS month_value,
                           SUM(total) AS total
                    FROM {$table}
                    {$where}
                    GROUP BY form_key, year_value, month_value
                    ORDER BY year_value, month_value
                ");
            }
        } else {
            $table = $wpdb->prefix . 'e_submissions';
            $name_expr = $display_mode === 'referer_title' || $display_mode === 'referer_title_distinct'
                ? "NULLIF(s.referer_title, '')"
                : "NULLIF(s.form_name, '')";
            $where = "WHERE s.status NOT LIKE '%trash%'";
            if (!empty($selected_keys)) {
                $placeholders = implode(', ', array_fill(0, count($selected_keys), '%s'));
                if ($display_mode === 'form_name' || $display_mode === 'referer_title') {
                    $where .= $wpdb->prepare(" AND CONCAT(s.element_id,'||',{$name_expr}) IN ($placeholders)", $selected_keys);
                } else {
                    $where .= $wpdb->prepare(" AND MD5(LOWER(TRIM({$name_expr}))) IN ($placeholders)", $selected_keys);
                }
            }
            $where .= mdp_get_email_exclusion_clause('s');
            if ($display_mode === 'form_name' || $display_mode === 'referer_title') {
                $rows = $wpdb->get_results("
                    SELECT CONCAT(s.element_id,'||',{$name_expr}) AS form_key,
                           {$name_expr} AS form_name,
                           YEAR(s.created_at_gmt) AS year_value,
                           MONTH(s.created_at_gmt) AS month_value,
                           COUNT(*) AS total
                    FROM {$table} s
                    {$where}
                    GROUP BY s.element_id, form_name, year_value, month_value
                    ORDER BY year_value, month_value
                ");
            } else {
                $rows = $wpdb->get_results("
                    SELECT MD5(LOWER(TRIM({$name_expr}))) AS form_key,
                           MIN({$name_expr}) AS form_name,
                           YEAR(s.created_at_gmt) AS year_value,
                           MONTH(s.created_at_gmt) AS month_value,
                           COUNT(*) AS total
                    FROM {$table} s
                    {$where}
                    GROUP BY form_key, year_value, month_value
                    ORDER BY year_value, month_value
                ");
            }
        }

        $forms = array();
        $form_titles = array();
        $year_totals = array();
        $years_map = array();

        foreach ($rows as $row) {
            if (empty($row->form_key)) {
                continue;
            }
            $form_key = sanitize_text_field($row->form_key);
            $year = (int) $row->year_value;
            $month = (int) $row->month_value;
            $total = (int) $row->total;
            if ($year <= 0 || $month <= 0) {
                continue;
            }
            $label = isset($row->form_name) ? sanitize_text_field($row->form_name) : '';
            if ($label === '') {
                $label = __('Unbenannte Seite', 'elementor-forms-statistics');
            }

            $years_map[$year] = true;
            if (!isset($year_totals[$year])) {
                $year_totals[$year] = $month_template;
            }
            if (!isset($year_totals[$year][$month])) {
                $year_totals[$year][$month] = 0;
            }
            $year_totals[$year][$month] += $total;

            if (!isset($forms[$form_key])) {
                $forms[$form_key] = array(
                    'title' => '',
                    'yearly' => array(),
                );
            }
            if (!isset($forms[$form_key]['yearly'][$year])) {
                $forms[$form_key]['yearly'][$year] = $month_template;
            }
            $forms[$form_key]['yearly'][$year][$month] += $total;

            if (!isset($form_titles[$form_key]) || $form_titles[$form_key] === '') {
                $form_titles[$form_key] = $label;
            }
        }

        foreach ($forms as $form_key => &$form_data) {
            foreach ($form_data['yearly'] as $year => &$months) {
                $months = array_replace($month_template, $months);
            }
            unset($months);
            $form_data['title'] = isset($form_titles[$form_key]) ? $form_titles[$form_key] : __('Unbenannte Seite', 'elementor-forms-statistics');
        }
        unset($form_data);

        $form_order = array_keys($forms);
        if (!empty($selected_keys)) {
            $selected_lookup = array_flip($selected_keys);
            $form_order = array_values(array_filter($form_order, function($form_key) use ($selected_lookup) {
                return isset($selected_lookup[$form_key]);
            }));
        }
        usort($form_order, function($a, $b) use ($forms) {
            return strcasecmp($forms[$a]['title'], $forms[$b]['title']);
        });

        $years_asc = array_keys($years_map);
        sort($years_asc);
        $years_desc = array_reverse($years_asc);

        foreach ($year_totals as $year => &$months) {
            $months = array_replace($month_template, $months);
        }
        unset($months);

        return array(
            'forms' => $forms,
            'form_order' => $form_order,
            'years_asc' => $years_asc,
            'years_desc' => $years_desc,
            'year_totals' => $year_totals,
            'current_year' => (int) date('Y'),
            'current_month' => (int) date('n'),
        );
    }

    $forms_indexed = $args['forms_indexed'];
    $title_overrides = mdp_get_form_title_overrides();
    $form_grouping = mdp_build_canonical_form_mapping($forms_indexed, $title_overrides, $display_mode);
    $canonical_forms = $form_grouping['canonical_forms'];
    $element_to_canonical = $form_grouping['element_to_canonical'];

    $canonical_selected_ids = array();
    foreach ($selected_ids as $form_id) {
        $mapped = isset($element_to_canonical[$form_id]) ? $element_to_canonical[$form_id] : $form_id;
        if ($mapped !== '') {
            $canonical_selected_ids[] = $mapped;
        }
    }
    $selected_ids = array_values(array_unique(array_filter($canonical_selected_ids)));
    $selected_member_ids = array();
    if (!empty($selected_ids)) {
        foreach ($selected_ids as $canonical_id) {
            if (isset($canonical_forms[$canonical_id]['member_ids'])) {
                $selected_member_ids = array_merge($selected_member_ids, $canonical_forms[$canonical_id]['member_ids']);
            } else {
                $selected_member_ids[] = $canonical_id;
            }
        }
        $selected_member_ids = array_values(array_unique(array_filter($selected_member_ids)));
    }

    $rows = array();
    if ($data_source === 'archive') {
        if (!mdp_archive_table_exists()) {
            $rows = array();
        } else {
            $table = mdp_get_archive_table_name();
            $where = "WHERE 1=1";
            if (!empty($selected_member_ids)) {
                $placeholders = implode(', ', array_fill(0, count($selected_member_ids), '%s'));
                $where .= $wpdb->prepare(" AND form_id IN ($placeholders)", $selected_member_ids);
            }
            $rows = $wpdb->get_results("
                SELECT form_id,
                       MAX(NULLIF(form_name, '')) AS form_title,
                       year AS year_value,
                       month AS month_value,
                       SUM(total) AS total
                FROM {$table}
                {$where}
                GROUP BY form_id, year_value, month_value
                ORDER BY year_value, month_value
            ");
        }
    } else {
        $table = $wpdb->prefix . 'e_submissions';
        $where = "WHERE s.status NOT LIKE '%trash%'";
        if (!empty($selected_member_ids)) {
            $placeholders = implode(', ', array_fill(0, count($selected_member_ids), '%s'));
            $where .= $wpdb->prepare(" AND s.element_id IN ($placeholders)", $selected_member_ids);
        }
        $where .= mdp_get_email_exclusion_clause('s');
        $rows = $wpdb->get_results("
            SELECT s.element_id AS form_id,
                   MAX(s.referer_title) AS form_title,
                   YEAR(s.created_at_gmt) AS year_value,
                   MONTH(s.created_at_gmt) AS month_value,
                   COUNT(*) AS total
            FROM {$table} s
            {$where}
            GROUP BY s.element_id, year_value, month_value
            ORDER BY year_value, month_value
        ");
    }

    $forms = array();
    $form_titles = array();
    $year_totals = array();
    $years_map = array();
    $month_template = mdp_get_empty_month_template();

    foreach ($rows as $row) {
        if (empty($row->form_id)) {
            continue;
        }
        $form_id = sanitize_text_field($row->form_id);
        $canonical_id = isset($element_to_canonical[$form_id]) ? $element_to_canonical[$form_id] : $form_id;
        $year = (int) $row->year_value;
        $month = (int) $row->month_value;
        $total = (int) $row->total;
        if ($year <= 0 || $month <= 0) {
            continue;
        }

        $years_map[$year] = true;
        if (!isset($year_totals[$year])) {
            $year_totals[$year] = $month_template;
        }
        if (!isset($year_totals[$year][$month])) {
            $year_totals[$year][$month] = 0;
        }
        $year_totals[$year][$month] += $total;

        if (!isset($forms[$canonical_id])) {
            $forms[$canonical_id] = array(
                'title' => '',
                'yearly' => array(),
            );
        }
        if (!isset($forms[$canonical_id]['yearly'][$year])) {
            $forms[$canonical_id]['yearly'][$year] = $month_template;
        }
        $forms[$canonical_id]['yearly'][$year][$month] += $total;

        if (!isset($form_titles[$canonical_id]) || $form_titles[$canonical_id] === '') {
            $form_titles[$canonical_id] = $row->form_title ? sanitize_text_field($row->form_title) : '';
        }
    }

    $selected_ids = array_unique($selected_ids);
    foreach ($selected_ids as $form_id) {
        if (!isset($forms[$form_id])) {
            $forms[$form_id] = array(
                'title' => '',
                'yearly' => array(),
            );
        }
    }

    foreach ($forms as $form_id => &$form_data) {
        foreach ($form_data['yearly'] as $year => &$months) {
            $months = array_replace($month_template, $months);
        }
        unset($months);
        $fallback = isset($canonical_forms[$form_id]['title'])
            ? $canonical_forms[$form_id]['title']
            : (isset($form_titles[$form_id]) ? $form_titles[$form_id] : '');
        $form_data['title'] = mdp_format_form_title($form_id, $forms_indexed, $fallback);
    }
    unset($form_data);

    $form_order = array_keys($forms);
    if (!empty($selected_ids)) {
        $selected_lookup = array_flip($selected_ids);
        $form_order = array_values(array_filter($form_order, function($form_id) use ($selected_lookup) {
            return isset($selected_lookup[$form_id]);
        }));
    }
    if ($display_mode === 'element_id') {
        usort($form_order, function($a, $b) {
            return strcasecmp($a, $b);
        });
    } else {
        usort($form_order, function($a, $b) use ($forms) {
            return strcasecmp($forms[$a]['title'], $forms[$b]['title']);
        });
    }

    $years_asc = array_keys($years_map);
    sort($years_asc);
    $years_desc = array_reverse($years_asc);

    foreach ($year_totals as $year => &$months) {
        $months = array_replace($month_template, $months);
    }
    unset($months);

return array(
        'forms' => $forms,
        'form_order' => $form_order,
        'years_asc' => $years_asc,
        'years_desc' => $years_desc,
        'year_totals' => $year_totals,
        'current_year' => (int) date('Y'),
        'current_month' => (int) date('n'),
    );
}

function mdp_get_schedule_interval() {
    $value = get_option('mdp_efs_email_interval', 'monthly');
    $valid = array('disabled', 'daily', 'weekly', 'monthly');
    if (!in_array($value, $valid, true)) {
        $value = 'monthly';
    }
    return $value;
}

function mdp_get_email_recipients() {
    $raw = get_option('mdp_efs_email_recipients', '');
    if (!is_string($raw) || $raw === '') {
        return [];
    }
    $parts = preg_split('/[\r\n,]+/', $raw);
    $emails = [];
    foreach ($parts as $part) {
        $trimmed = trim($part);
        if ($trimmed === '') {
            continue;
        }
        $email = sanitize_email($trimmed);
        if ($email) {
            $emails[] = strtolower($email);
        }
    }
    return array_values(array_unique($emails));
}

function mdp_get_email_recipients_text() {
    $emails = mdp_get_email_recipients();
    return implode("\n", $emails);
}

function mdp_get_default_email_message() {
    return __("Hallo,

anbei die aktuelle Übersicht der Formular-Anfragen.
%stats_link%

Deine freundliche Website
%s

", 'elementor-forms-statistics');
}

function mdp_get_email_message() {
    $default = mdp_get_default_email_message();
    $message = get_option('mdp_efs_email_message', $default);
    if (!is_string($message) || $message === '') {
        $message = $default;
    }
    return $message;
}

function mdp_get_email_subject_template() {
    $template = get_option('mdp_efs_email_subject', '📈 ' . __('Anfragen Statistik – %s', 'elementor-forms-statistics'));
    if (!is_string($template) || $template === '') {
        $template = '📈 ' . __('Anfragen Statistik – %s', 'elementor-forms-statistics');
    }
    return $template;
}

function mdp_should_include_stats_attachment() {
    $value = get_option('mdp_efs_include_attachment', '1');
    return $value !== '0';
}

function mdp_build_email_subject($date_string = '') {
    $template = mdp_get_email_subject_template();
    $formatted_date = $date_string !== '' ? $date_string : date_i18n('F Y');
    if (strpos($template, '%s') !== false) {
        return sprintf($template, $formatted_date);
    }
    return trim($template . ' ' . $formatted_date);
}

function mdp_build_styled_email_body($message_with_link, $link_note, $export_url, $link_label) {
    $message_html = wpautop($message_with_link);
    $badge = __('Formular-Update', 'elementor-forms-statistics');

    $note_block = $link_note ? '<div style="margin:24px 0 0;padding:0;">' . $link_note . '</div>' : '';
    $branding_url = 'https://www.medienproduktion.biz/elementor-forms-statistics/';
    $plugin_version = mdp_get_plugin_version_string();
    $brand_text = sprintf(__('%s · 2025 · v%s', 'elementor-forms-statistics'), __('Elementor Forms Statistics', 'elementor-forms-statistics'), $plugin_version);
    $branding_block = '<p style="margin:20px 0 0;font-size:11px;color:#9ca3b4;text-align:center;">'
        . '<a href="' . esc_url($branding_url) . '" target="_blank" rel="noopener noreferrer" style="color:#9ca3b4;text-decoration:none;">'
        . esc_html($brand_text)
        . '</a></p>';
    return '
    <div style="background:#f5f6fb;padding:48px 16px;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica Neue,Arial,sans-serif;color:#1f2a44;">
        <div style="max-width:720px;margin:0 auto;background:#ffffff;border-radius:20px;box-shadow:0 25px 60px rgba(31,42,68,0.12);padding:40px 48px;">
            <p style="margin:0 0 6px;font-size:12px;font-weight:600;letter-spacing:0.2em;text-transform:uppercase;color:#8c96b0;">' . esc_html($badge) . '</p>
            <div style="font-size:16px;line-height:1.7;color:#2d3756;">' . $message_html . '</div>
            ' . $note_block . '
            ' . $branding_block . '
        </div>
    </div>';
}

function mdp_get_default_curve_color_slots() {
    return array(
        array('color' => '#a4bbb0', 'alpha' => 0.95),
        array('color' => '#c9b1a5', 'alpha' => 0.9),
        array('color' => '#d49484', 'alpha' => 0.88),
        array('color' => '#edc1a6', 'alpha' => 0.85),
        array('color' => '#bec0b2', 'alpha' => 0.82),
        array('color' => '#d7b69f', 'alpha' => 0.8),
        array('color' => '#8fa6bf', 'alpha' => 0.77),
        array('color' => '#9dc1c8', 'alpha' => 0.74),
        array('color' => '#c6d5dd', 'alpha' => 0.7),
        array('color' => '#d7d6d1', 'alpha' => 0.66),
    );
}

function mdp_get_curve_color_slots() {
    $stored = get_option('mdp_efs_curve_color_slots');
    $defaults = mdp_get_default_curve_color_slots();
    $slots = array();
    for ($i = 0; $i < 10; $i++) {
        $slot = isset($stored[$i]) && is_array($stored[$i]) ? $stored[$i] : array();
        $color = isset($slot['color']) ? sanitize_hex_color($slot['color']) : '';
        $alpha = isset($slot['alpha']) ? (float) $slot['alpha'] : $defaults[$i]['alpha'];
        if ($color === '') {
            $color = $defaults[$i]['color'];
        }
        if ($alpha < 0) {
            $alpha = 0;
        } elseif ($alpha > 1) {
            $alpha = 1;
        }
        $slots[] = array(
            'color' => $color,
            'alpha' => $alpha,
        );
    }
    return $slots;
}

function mdp_hex_to_rgba($hex, $alpha = 1.0) {
    $hex = sanitize_hex_color($hex);
    if (!$hex) {
        $hex = '#000000';
    }
    $hex = ltrim($hex, '#');
    if (strlen($hex) === 3) {
        $hex = $hex[0] . $hex[0] . $hex[1] . $hex[1] . $hex[2] . $hex[2];
    }
    $r = hexdec(substr($hex, 0, 2));
    $g = hexdec(substr($hex, 2, 2));
    $b = hexdec(substr($hex, 4, 2));
    return 'rgba(' . $r . ',' . $g . ',' . $b . ',' . $alpha . ')';
}

function mdp_get_email_schedule_description($interval) {
    $descriptions = array(
        'disabled' => __('Deaktiviert', 'elementor-forms-statistics'),
        'daily' => __('Täglich', 'elementor-forms-statistics'),
        'weekly' => __('Wöchentlich', 'elementor-forms-statistics'),
        'monthly' => __('Monatlich', 'elementor-forms-statistics'),
    );
    return isset($descriptions[$interval]) ? $descriptions[$interval] : $interval;
}

function mdp_get_email_send_time() {
    $value = get_option('mdp_efs_email_time', '08:00');
    if (!is_string($value) || !preg_match('/^([01]\d|2[0-3]):([0-5]\d)$/', $value)) {
        $value = '08:00';
    }
    return $value;
}

function mdp_get_email_day_of_month() {
    $value = (int) get_option('mdp_efs_email_day_of_month', 1);
    if ($value < 1) {
        $value = 1;
    } elseif ($value > 31) {
        $value = 31;
    }
    return $value;
}

function mdp_get_email_weekday_choices() {
    return array(
        'monday' => __('Montag', 'elementor-forms-statistics'),
        'tuesday' => __('Dienstag', 'elementor-forms-statistics'),
        'wednesday' => __('Mittwoch', 'elementor-forms-statistics'),
        'thursday' => __('Donnerstag', 'elementor-forms-statistics'),
        'friday' => __('Freitag', 'elementor-forms-statistics'),
        'saturday' => __('Samstag', 'elementor-forms-statistics'),
        'sunday' => __('Sonntag', 'elementor-forms-statistics'),
    );
}

function mdp_get_email_weekday() {
    $value = strtolower(get_option('mdp_efs_email_weekday', 'monday'));
    $choices = mdp_get_email_weekday_choices();
    if (!isset($choices[$value])) {
        $value = 'monday';
    }
    return $value;
}

function mdp_get_email_time_parts() {
    $time = mdp_get_email_send_time();
    $parts = explode(':', $time);
    $hour = isset($parts[0]) ? (int) $parts[0] : 8;
    $minute = isset($parts[1]) ? (int) $parts[1] : 0;
    return array($hour, $minute);
}

function mdp_calculate_next_email_timestamp($interval) {
    if ($interval === 'disabled') {
        return 0;
    }

    $timezone = wp_timezone();
    $now = new DateTime('now', $timezone);
    list($hour, $minute) = mdp_get_email_time_parts();

    if ($interval === 'daily') {
        $target = clone $now;
        $target->setTime($hour, $minute, 0);
        if ($target <= $now) {
            $target->modify('+1 day');
        }
        return $target->getTimestamp();
    }

    if ($interval === 'weekly') {
        $weekday = mdp_get_email_weekday();
        $target = clone $now;
        $target->setTime($hour, $minute, 0);
        $current_weekday = strtolower($now->format('l'));
        if ($weekday === $current_weekday && $target > $now) {
            return $target->getTimestamp();
        }
        $target = new DateTime('next ' . $weekday, $timezone);
        $target->setTime($hour, $minute, 0);
        return $target->getTimestamp();
    }

    if ($interval === 'monthly') {
        $day = mdp_get_email_day_of_month();
        $target = clone $now;
        $target_day = min($day, (int) $target->format('t'));
        $target->setDate((int) $target->format('Y'), (int) $target->format('m'), $target_day);
        $target->setTime($hour, $minute, 0);
        if ($target <= $now) {
            $target->modify('first day of next month');
            $target_day = min($day, (int) $target->format('t'));
            $target->setDate((int) $target->format('Y'), (int) $target->format('m'), $target_day);
            $target->setTime($hour, $minute, 0);
        }
        return $target->getTimestamp();
    }

    return time() + MINUTE_IN_SECONDS;
}

function mdp_get_cron_schedule_slug($interval) {
    switch ($interval) {
        case 'daily':
            return 'daily';
        case 'weekly':
            return 'weekly';
        case 'monthly':
            return 'mdp_monthly';
        default:
            return '';
    }
}

function mdp_add_custom_cron_schedules($schedules) {
    if (!isset($schedules['mdp_monthly'])) {
        $schedules['mdp_monthly'] = array(
            'interval' => DAY_IN_SECONDS * 30,
            'display'  => __('Einmal im Monat', 'elementor-forms-statistics'),
        );
    }
    return $schedules;
}

function mdp_maybe_schedule_stats_email() {
    $interval = mdp_get_schedule_interval();
    $recipients = mdp_get_email_recipients();
    $hook = 'mdp_send_stats_email';
    $schedule_mode = get_option('mdp_efs_schedule_mode', '');
    if ($schedule_mode !== 'single') {
        wp_clear_scheduled_hook($hook);
        update_option('mdp_efs_schedule_mode', 'single');
    }

    if ($interval === 'disabled' || empty($recipients)) {
        wp_clear_scheduled_hook($hook);
        return;
    }

    if (wp_next_scheduled($hook)) {
        return;
    }

    $next_timestamp = mdp_calculate_next_email_timestamp($interval);
    if ($next_timestamp > 0) {
        wp_schedule_single_event($next_timestamp, $hook);
    }
}

function mdp_reset_stats_email_schedule() {
    wp_clear_scheduled_hook('mdp_send_stats_email');
    mdp_maybe_schedule_stats_email();
}

function mdp_get_submission_cleanup_interval() {
    $valid = array('disabled', '1h', '1d', '1m', '1y');
    $interval = get_option('mdp_efs_submission_cleanup_interval', 'disabled');
    return in_array($interval, $valid, true) ? $interval : 'disabled';
}

function mdp_get_submission_cleanup_seconds() {
    switch (mdp_get_submission_cleanup_interval()) {
        case '1h':
            return HOUR_IN_SECONDS;
        case '1d':
            return DAY_IN_SECONDS;
        case '1m':
            return DAY_IN_SECONDS * 30;
        case '1y':
            return DAY_IN_SECONDS * 365;
        default:
            return 0;
    }
}

function mdp_clear_elementor_submissions_callback() {
    global $wpdb;
    $seconds = mdp_get_submission_cleanup_seconds();
    if ($seconds <= 0) {
        wp_clear_scheduled_hook('mdp_clear_elementor_submissions');
        return;
    }
    $threshold = gmdate('Y-m-d H:i:s', current_time('timestamp', true) - $seconds);
    $status_like = '%trash%';
    $condition_with_alias = $wpdb->prepare('s.status NOT LIKE %s AND s.created_at_gmt <= %s', $status_like, $threshold);
    $condition_plain = $wpdb->prepare('status NOT LIKE %s AND created_at_gmt <= %s', $status_like, $threshold);
    $values_table = $wpdb->prefix . 'e_submissions_values';
    $submissions_table = $wpdb->prefix . 'e_submissions';
    $wpdb->query("DELETE ev FROM {$values_table} ev INNER JOIN {$submissions_table} s ON ev.submission_id = s.ID WHERE {$condition_with_alias}");
    $wpdb->query("DELETE FROM {$submissions_table} WHERE {$condition_plain}");
    mdp_schedule_submission_cleanup(true);
}

function mdp_schedule_submission_cleanup($force_reschedule = false) {
    $seconds = mdp_get_submission_cleanup_seconds();
    if ($seconds <= 0) {
        wp_clear_scheduled_hook('mdp_clear_elementor_submissions');
        return;
    }
    if ($force_reschedule) {
        wp_clear_scheduled_hook('mdp_clear_elementor_submissions');
    }
    if (!wp_next_scheduled('mdp_clear_elementor_submissions')) {
        wp_schedule_single_event(time() + $seconds, 'mdp_clear_elementor_submissions');
    }
}

add_action('mdp_clear_elementor_submissions', 'mdp_clear_elementor_submissions_callback');

function mdp_send_stats_email_callback($manual = false) {
    $recipients = mdp_get_email_recipients();
    if (empty($recipients)) {
        if (!$manual) {
            wp_clear_scheduled_hook('mdp_send_stats_email');
        }
        return false;
    }

    $include_attachment = mdp_should_include_stats_attachment();

    $subject = mdp_build_email_subject(date_i18n('F Y'));
    $message = mdp_get_email_message();
    $site_url_text = esc_url(trailingslashit(home_url()));
    $email_font_stack = "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', 'Segoe UI', 'Open Sans', 'Helvetica Neue', sans-serif";
    $blog_name = wp_specialchars_decode(get_option('blogname'), ENT_QUOTES);
    $headers = array(
        'Content-Type: text/html; charset=' . get_option('blog_charset'),
        'From: ' . $blog_name . ' <' . get_option('admin_email') . '>',
    );

    $html_export = custom_menu_callback(true, array(
        'export' => true,
        'inline_styles' => true,
        'include_html_document' => true,
        'include_chartjs' => true,
    ));

    $export_url = mdp_save_stats_export_file($html_export);
    $link_note = '';
    $link_label = mdp_get_export_link_label();
    $button_html = mdp_build_stats_link_button($export_url, $link_label);
    $replacements = array(
        '%s' => $site_url_text,
        '%stats_link%' => $button_html,
        'stats_link%' => $button_html,
        '%stats_link' => $button_html,
    );
    $message_with_link = strtr($message, $replacements);
    if ($export_url) {
        $message_with_link = preg_replace(
            '/https?:\\/\\/[\\w\\-\\.\\/%]+stats_link%/i',
            $button_html,
            $message_with_link
        );
    }

    $body = mdp_build_styled_email_body($message_with_link, $link_note, $export_url, $link_label);

    $attachments = array();
    if ($include_attachment) {
        $temp_dir = trailingslashit(get_temp_dir());
        $attachment_filename = 'Anfragen Statistik ' . date_i18n('Y-m') . '.html';
        $attachment_basename = wp_unique_filename($temp_dir, $attachment_filename);
        $attachment_path = $temp_dir . $attachment_basename;
        if (file_put_contents($attachment_path, $html_export) !== false) {
            $attachments[] = $attachment_path;
        }
    }

    $sent = wp_mail($recipients, $subject, $body, $headers, $attachments);

    if (!empty($attachments)) {
        foreach ($attachments as $attachment) {
            if (file_exists($attachment)) {
                unlink($attachment);
            }
        }
    }

    if (!$manual) {
        wp_clear_scheduled_hook('mdp_send_stats_email');
        mdp_maybe_schedule_stats_email();
    }

    return (bool) $sent;
}

function mdp_get_stats_export_folder_slug() {
    $option = get_option('mdp_efs_export_folder', 'mdpstats');
    $slug = sanitize_title($option);
    if ($slug === '') {
        $slug = 'elementor-forms-statistics';
    }
    return $slug;
}

function mdp_get_stats_export_subdir() {
    return mdp_get_stats_export_folder_slug();
}

function mdp_get_stats_export_dir() {
    $uploads = wp_upload_dir();
    return trailingslashit($uploads['basedir']) . mdp_get_stats_export_subdir();
}

function mdp_get_stats_export_url() {
    $uploads = wp_upload_dir();
    return trailingslashit($uploads['baseurl']) . mdp_get_stats_export_subdir();
}

function mdp_ensure_stats_export_dir() {
    $dir = mdp_get_stats_export_dir();
    if (!wp_mkdir_p($dir)) {
        return false;
    }
    mdp_ensure_stats_export_htaccess($dir);
    return $dir;
}

function mdp_ensure_stats_export_htaccess($dir) {
    $htaccess_path = trailingslashit($dir) . '.htaccess';
    $lines = array(
        '# Generated by MDP Elementor Forms Extras',
        '# Der Ordner ist nur über Token-Links erreichbar.',
        '<IfModule mod_headers.c>',
        'Header set X-Robots-Tag "noindex, nofollow, noarchive"',
        '</IfModule>',
        'Options -Indexes',
    );
    file_put_contents($htaccess_path, implode("\n", $lines) . "\n", LOCK_EX);
}

function mdp_remove_directory_tree($dir) {
    if (!$dir || !is_dir($dir)) {
        return;
    }
    $items = scandir($dir);
    foreach ($items as $item) {
        if ($item === '.' || $item === '..') {
            continue;
        }
        $path = trailingslashit($dir) . $item;
        if (is_dir($path)) {
            mdp_remove_directory_tree($path);
        } else {
            @unlink($path);
        }
    }
    @rmdir($dir);
}

function mdp_build_stats_link_button($export_url, $link_label) {
    if (!$export_url) {
        return '';
    }
    $label = $link_label !== '' ? $link_label : __('Statistik öffnen', 'elementor-forms-statistics');
    $url = esc_url($export_url);
    return '<a href="' . $url . '" target="_blank" rel="noopener noreferrer" style="display:inline-block;margin:16px auto 0;padding:14px 32px;background:#1f2a44;color:#ffffff;font-weight:600;text-decoration:none;border-radius:999px;box-shadow:0 10px 30px rgba(31,42,68,0.25);">' . esc_html($label) . '</a>';
}

function mdp_uninstall_plugin() {
    if (get_option('mdp_efs_clean_on_uninstall', '0') !== '1') {
        return;
    }

    global $wpdb;
    wp_clear_scheduled_hook('mdp_send_stats_email');
    wp_clear_scheduled_hook('mdp_clear_elementor_submissions');
    $table = mdp_get_archive_table_name();
    $wpdb->query("DROP TABLE IF EXISTS {$table}");

    $like_prefix = $wpdb->esc_like('mdp_efs_') . '%';
    $wpdb->query($wpdb->prepare("DELETE FROM {$wpdb->options} WHERE option_name LIKE %s", $like_prefix));

    mdp_remove_directory_tree(mdp_get_stats_export_dir());
}

function mdp_save_stats_export_file($html_export) {
    $dir = mdp_ensure_stats_export_dir();
    if (!$dir || $html_export === '') {
        return '';
    }
    $timestamp = date_i18n('Y-m-d_H-i-s');
    $random = substr(wp_generate_password(8, false, false), 0, 8);
    $base_name = 'anfragen-statistik-' . $timestamp . '-' . $random;
    $filename = wp_unique_filename($dir, sanitize_file_name($base_name . '.html'));
    $path = trailingslashit($dir) . $filename;
    if (file_put_contents($path, $html_export) === false) {
        return '';
    }
    return trailingslashit(mdp_get_stats_export_url()) . rawurlencode($filename);
}

function mdp_get_stats_link_placeholder() {
    return '%stats_link%';
}

function mdp_get_export_link_label() {
    $default = __('Statistik öffnen', 'elementor-forms-statistics');
    $label = get_option('mdp_efs_export_link_label', $default);
    if (!is_string($label) || $label === '') {
        return $default;
    }
    return $label;
}

function mdp_send_stats_now_handler() {
    if (!current_user_can('edit_posts')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }

    if (!isset($_POST['mdp_send_now_nonce']) || !wp_verify_nonce($_POST['mdp_send_now_nonce'], 'mdp_send_stats_now')) {
        wp_die(__('Ungültige Anfrage.', 'elementor-forms-statistics'));
    }

    $result = mdp_send_stats_email_callback(true);
    $status = $result ? 'success' : 'error';
    $redirect = wp_get_referer();
    if (!$redirect) {
        $redirect = admin_url('admin.php?page=statistiken-einstellungen');
    }
    wp_safe_redirect(add_query_arg('mdp_send_status', $status, $redirect));
    exit;
}

function mdp_get_font_base64($absolute_path) {
    if (!file_exists($absolute_path)) {
        return '';
    }
    $contents = file_get_contents($absolute_path);
    if ($contents === false) {
        return '';
    }
    return base64_encode($contents);
}

function mdp_get_inline_export_styles() {
    $css = '';
    $css_path = plugin_dir_path(__FILE__) . 'assets/css/style.css';
    if (file_exists($css_path)) {
        $file_css = file_get_contents($css_path);
        if ($file_css !== false) {
            $css .= $file_css;
        }
    }

    $dashicons_path = trailingslashit(ABSPATH) . 'wp-includes/fonts/dashicons.woff2';
    $dashicons_base64 = mdp_get_font_base64($dashicons_path);
    if ($dashicons_base64) {
        $css .= "@font-face{font-family:'dashicons';font-style:normal;font-weight:normal;font-display:swap;src:url(data:font/woff2;base64," . $dashicons_base64 . ") format('woff2');}";
    }

    $font_stack = "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', 'Segoe UI', 'Open Sans', 'Helvetica Neue', sans-serif";
    $css .= "
    body.mdp-export-body,
    .mdp-email-container {
        font-family: {$font_stack};
        background: #e7eaef;
        color: #1d2327;
        margin: 0;
        padding: 32px 0 48px;
    }
    .mdp-email-container {
        width: 100%;
    }
    body.mdp-export-body .wrap,
    .mdp-email-container .wrap {
        max-width: 1180px;
        margin: 0 auto;
        background: transparent;
        padding: 0 40px;
        border-radius: 0;
        box-shadow: none;
        border: 0;
        font-family: {$font_stack};
    }
    .mdp-email-message {
        font-family: {$font_stack};
        max-width: 1180px;
        margin: 0 auto 24px;
        padding: 0 40px;
        color: #1d2327;
        line-height: 1.5;
        font-size: 14px;
    }
    .mdp-email-container .wrap.mdp-stats-root,
    .mdp-email-container h1,
    .mdp-email-container h2,
    .mdp-email-container .mdp-version,
    .mdp-email-container table,
    .mdp-email-container .mdp-chart,
    .mdp-email-container .mdp-table,
    .mdp-email-container .checkbox-container,
    .mdp-email-container p {
        font-family: {$font_stack};
    }
    body.mdp-export-body .mdp-chart,
    body.mdp-export-body .mdp-table,
    .mdp-email-container .mdp-chart,
    .mdp-email-container .mdp-table {
        width: 100%;
        max-width: 1090px;
        margin-left: auto;
        margin-right: auto;
    }
    .mdp-email-container .mdp-year-heading,
    body.mdp-export-body .mdp-year-heading {
        max-width: 1090px;
        margin: 40px auto 10px;
        font-size: 18px;
        font-weight: 600;
        color: #1d2327;
    }
    body.mdp-export-body .checkbox-container,
    .mdp-email-container .checkbox-container {
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    }
    body.mdp-export-body .mdp-chart,
    .mdp-email-container .mdp-chart {
        height: 520px;
    }
    body.mdp-export-body .mdp-version,
    .mdp-email-container .mdp-version {
        text-align: center;
        color: #50575e;
    }
    body.mdp-export-body h1,
    body.mdp-export-body h2,
    .mdp-email-container h1,
    .mdp-email-container h2 {
        color: #1d2327;
    }
    ";

    return $css;
}

function mdp_get_emphasis_threshold() {
    $value = get_option('mdp_efs_emphasis_threshold', 5);
    $value = is_numeric($value) ? floatval($value) : 5;
    if ($value < 0) {
        $value = 0;
    }
    return $value;
}

/**
 * Calculate an inline background style comparing current month values with previous year.
 */
function mdp_get_comparison_background($current_value, $previous_value, $has_previous_year_data) {
    if (!$has_previous_year_data) {
        return '';
    }

    $current_value = (int) $current_value;
    $previous_value = (int) $previous_value;

    if ($current_value === $previous_value) {
        return '';
    }

    $delta = $current_value - $previous_value;
    $threshold = mdp_get_emphasis_threshold();
    if ($threshold > 0 && abs($delta) < $threshold) {
        return '';
    }
    $ratio = 0.0;

    if ($previous_value === 0) {
        if ($current_value === 0) {
            return '';
        }
        $ratio = min($current_value / 5, 1); // normalize growth when no previous data
    } else {
        $ratio = min(abs($delta) / max($previous_value, 1), 1);
    }

    $alpha = 0.03 + (0.12 * $ratio); // subtle pastel intensity
    $color = $delta > 0
        ? 'rgba(76, 175, 80,' . $alpha . ')'   // green tones
        : 'rgba(244, 67, 54,' . $alpha . ')';  // red tones

    return ' style="background-color:' . esc_attr($color) . ';"';
}

function mdp_should_emphasize_change($current_value, $previous_value, $has_previous_year_data) {
    if (!$has_previous_year_data) {
        return false;
    }

    $current_value = (int) $current_value;
    $previous_value = (int) $previous_value;

    if ($current_value === $previous_value) {
        return false;
    }

    $delta = abs($current_value - $previous_value);
    $threshold = mdp_get_emphasis_threshold();
    return $delta >= $threshold;
}

function mdp_get_excluded_emails() {
    $raw = get_option('mdp_efs_excluded_emails', '');
    if (!is_string($raw) || $raw === '') {
        return [];
    }
    $lines = preg_split('/\r\n|\r|\n/', $raw);
    $patterns = [];
    foreach ($lines as $line) {
        $trimmed = trim($line);
        if ($trimmed === '') {
            continue;
        }
        $patterns[] = strtolower(sanitize_text_field($trimmed));
    }
    return array_values(array_unique(array_filter($patterns)));
}

function mdp_get_excluded_emails_text() {
    $emails = mdp_get_excluded_emails();
    return implode("\n", $emails);
}

function mdp_prepare_email_pattern_for_like($pattern) {
    $pattern = strtolower(trim($pattern));
    if ($pattern === '') {
        return '';
    }
    $pattern = str_replace(array('*', '?'), array('%', '_'), $pattern);
    if (strpos($pattern, '%') === false && strpos($pattern, '_') === false) {
        $pattern = '%' . $pattern . '%';
    }
    return $pattern;
}

function mdp_get_email_exclusion_clause($table_alias) {
    $emails = mdp_get_excluded_emails();
    if (empty($emails)) {
        return '';
    }

    $like_conditions = [];
    $params = array('%email%');

    foreach ($emails as $pattern) {
        $like_pattern = mdp_prepare_email_pattern_for_like($pattern);
        if ($like_pattern === '') {
            continue;
        }
        $like_conditions[] = 'LOWER(ev.value) LIKE %s';
        $params[] = $like_pattern;
    }

    if (empty($like_conditions)) {
        return '';
    }

    global $wpdb;
    $condition_sql = implode(' OR ', $like_conditions);

    return $wpdb->prepare(
        " AND NOT EXISTS (
            SELECT 1 FROM " . $wpdb->prefix . "e_submissions_values ev
            WHERE ev.submission_id = {$table_alias}.ID
            AND LOWER(ev.`key`) LIKE %s
            AND ({$condition_sql})
        )",
        $params
    );
}

function mdp_settings_page_callback() {
    if (!mdp_user_can_access_menu('statistiken-einstellungen')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }

    $message = '';
    if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['mdp_settings_nonce']) && wp_verify_nonce($_POST['mdp_settings_nonce'], 'mdp_save_settings')) {
        $emails_raw = isset($_POST['excluded_emails']) ? wp_unslash($_POST['excluded_emails']) : '';
        $lines = preg_split('/\r\n|\r|\n/', $emails_raw);
        $sanitized = [];
        foreach ($lines as $line) {
            $trimmed = trim($line);
            if ($trimmed === '') {
                continue;
            }
            $pattern = sanitize_text_field($trimmed);
            if ($pattern !== '') {
                $sanitized[] = strtolower($pattern);
            }
        }
        $threshold_value = isset($_POST['emphasis_threshold']) ? floatval($_POST['emphasis_threshold']) : 5;
        if ($threshold_value < 0) {
            $threshold_value = 0;
        }
        $export_folder_slug = isset($_POST['export_folder']) ? sanitize_title($_POST['export_folder']) : '';
        if ($export_folder_slug === '') {
            $export_folder_slug = 'mdpstats';
        }
        $clean_on_uninstall = isset($_POST['clean_on_uninstall']) ? '1' : '0';
        $cleanup_interval = isset($_POST['submission_cleanup_interval']) ? sanitize_text_field($_POST['submission_cleanup_interval']) : mdp_get_submission_cleanup_interval();
        $valid_intervals = array('disabled', '1h', '1d', '1m', '1y');
        if (!in_array($cleanup_interval, $valid_intervals, true)) {
            $cleanup_interval = 'disabled';
        }

        $menu_roles = array();
        $valid_roles = array_keys(mdp_get_menu_role_choices());
        $menu_items = array_keys(mdp_get_menu_items());
        if (isset($_POST['menu_roles']) && is_array($_POST['menu_roles'])) {
            foreach ($_POST['menu_roles'] as $menu_key => $roles) {
                if (!in_array($menu_key, $menu_items, true) || !is_array($roles)) {
                    continue;
                }
                $menu_roles[$menu_key] = array();
                foreach ($roles as $role) {
                    $role = sanitize_key($role);
                    if ($role !== '' && in_array($role, $valid_roles, true)) {
                        $menu_roles[$menu_key][] = $role;
                    }
                }
                $menu_roles[$menu_key] = array_values(array_unique($menu_roles[$menu_key]));
            }
        }

        $curve_color_slots = array();
        $default_curve_slots = mdp_get_default_curve_color_slots();
        for ($i = 0; $i < 10; $i++) {
            $color_field = isset($_POST['curve_color_' . $i]) ? sanitize_hex_color($_POST['curve_color_' . $i]) : '';
            $alpha_field = isset($_POST['curve_alpha_' . $i]) ? floatval($_POST['curve_alpha_' . $i]) : round($default_curve_slots[$i]['alpha'] * 100);
            $alpha_value = max(0, min(100, $alpha_field)) / 100;
            if ($color_field === '') {
                $color_field = $default_curve_slots[$i]['color'];
            }
            $curve_color_slots[] = array(
                'color' => $color_field,
                'alpha' => $alpha_value,
            );
        }

        $display_mode = isset($_POST['form_display_mode']) ? sanitize_text_field($_POST['form_display_mode']) : '';
        if (!in_array($display_mode, array('element_id', 'form_name', 'form_name_distinct', 'referer_title', 'referer_title_distinct'), true)) {
            $display_mode = 'element_id';
        }
        $selected_for_mode = isset($_POST['selected_form_ids']) && is_array($_POST['selected_form_ids'])
            ? array_map('sanitize_text_field', $_POST['selected_form_ids'])
            : [];
        $stored_selections = get_option('mdp_efs_selected_form_ids', []);
        if (!is_array($stored_selections) || !isset($stored_selections['element_id']) && !isset($stored_selections['form_name']) && !isset($stored_selections['form_name_distinct']) && !isset($stored_selections['referer_title']) && !isset($stored_selections['referer_title_distinct'])) {
            $stored_selections = array(
                'element_id' => is_array($stored_selections) ? array_map('sanitize_text_field', $stored_selections) : [],
                'form_name' => [],
                'form_name_distinct' => [],
                'referer_title' => [],
                'referer_title_distinct' => [],
            );
        }
        $stored_selections[$display_mode] = array_values(array_filter($selected_for_mode));
        update_option('mdp_efs_selected_form_ids', $stored_selections);
        $form_title_overrides = [];
        if (isset($_POST['selected_form_titles']) && is_array($_POST['selected_form_titles'])) {
            foreach ($_POST['selected_form_titles'] as $canonical_id => $title_value) {
                $canonical_id = sanitize_text_field($canonical_id);
                if ($canonical_id === '') {
                    continue;
                }
                $title_value = is_string($title_value) ? sanitize_text_field($title_value) : '';
                if ($title_value === '') {
                    continue;
                }
                $form_title_overrides[$canonical_id] = $title_value;
            }
        }
        update_option('mdp_efs_form_display_mode', $display_mode);
        update_option('mdp_efs_form_title_overrides', $form_title_overrides);
        update_option('mdp_efs_excluded_emails', implode("\n", $sanitized));
        update_option('mdp_efs_emphasis_threshold', $threshold_value);
        update_option('mdp_efs_export_folder', $export_folder_slug);
        update_option('mdp_efs_curve_color_slots', $curve_color_slots);
        update_option('mdp_efs_clean_on_uninstall', $clean_on_uninstall);
        update_option('mdp_efs_submission_cleanup_interval', $cleanup_interval);
        update_option('mdp_efs_menu_roles', $menu_roles);
        mdp_ensure_stats_export_dir();
        $message = __('Einstellungen gespeichert.', 'elementor-forms-statistics');
        mdp_schedule_submission_cleanup(true);
    }

    $stored_text = get_option('mdp_efs_excluded_emails', '');
    $display_mode = mdp_get_form_display_mode();
    $selected_form_ids = mdp_get_selected_form_ids($display_mode);
    $form_rows = mdp_get_archive_form_rows($display_mode);
    $form_order = array_keys($form_rows);
    if ($display_mode === 'form_name' || $display_mode === 'form_name_distinct' || $display_mode === 'referer_title' || $display_mode === 'referer_title_distinct') {
        usort($form_order, function($a, $b) use ($form_rows) {
            return strcasecmp($form_rows[$a]['title'], $form_rows[$b]['title']);
        });
    } else {
        sort($form_order, SORT_STRING);
    }
    $emphasis_threshold = mdp_get_emphasis_threshold();
    $export_folder_slug = mdp_get_stats_export_folder_slug();
    $curve_color_slots = mdp_get_curve_color_slots();
    $clean_on_uninstall = get_option('mdp_efs_clean_on_uninstall', '0');
    $cleanup_interval = mdp_get_submission_cleanup_interval();
    $menu_roles = mdp_get_menu_roles();
    $menu_role_choices = mdp_get_menu_role_choices();
    $menu_items = mdp_get_menu_items();
    ?>
    <div class="wrap mdp-stats-root">
        <h1><?php _e('Anfragen Einstellungen', 'elementor-forms-statistics'); ?></h1>
        <?php if ($message) : ?>
            <div class="notice notice-success"><p><?php echo esc_html($message); ?></p></div>
        <?php endif; ?>
        <form method="post">
            <?php wp_nonce_field('mdp_save_settings', 'mdp_settings_nonce'); ?>
            <div class="mdp-form-wrapper">
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label><?php _e('Sichtbare Formulare in Statistik', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <div class="mdp-form-field-inline">
                                <label for="form_display_mode"><?php _e('Liste nach', 'elementor-forms-statistics'); ?></label>
                                <select name="form_display_mode" id="form_display_mode">
                                    <option value="element_id" <?php selected($display_mode, 'element_id'); ?>><?php _e('Element ID', 'elementor-forms-statistics'); ?></option>
                                    <option value="form_name" <?php selected($display_mode, 'form_name'); ?>><?php _e('Formularname', 'elementor-forms-statistics'); ?></option>
                                    <option value="form_name_distinct" <?php selected($display_mode, 'form_name_distinct'); ?>><?php _e('Formularname (Distinct)', 'elementor-forms-statistics'); ?></option>
                                    <option value="referer_title" <?php selected($display_mode, 'referer_title'); ?>><?php _e('Seiten-Referer', 'elementor-forms-statistics'); ?></option>
                                    <option value="referer_title_distinct" <?php selected($display_mode, 'referer_title_distinct'); ?>><?php _e('Seiten-Referer (Distinct)', 'elementor-forms-statistics'); ?></option>
                                </select>
                            </div>
                            <?php
                            $has_forms_to_show = !empty($form_order);
                            if (!$has_forms_to_show) :
                            ?>
                                <p><?php _e('Keine Seiten gefunden.', 'elementor-forms-statistics'); ?></p>
                            <?php else : ?>
                                <div class="mdp-form-field-table">
                                    <table class="widefat striped">
                                        <thead>
                                            <tr>
                                                <th scope="col" class="check-column"><span class="screen-reader-text"><?php _e('Auswählen', 'elementor-forms-statistics'); ?></span></th>
                                                <th scope="col"><?php _e('ID', 'elementor-forms-statistics'); ?></th>
                                                <th scope="col"><?php _e('Formular', 'elementor-forms-statistics'); ?></th>
                                                <th scope="col"><?php _e('Einträge', 'elementor-forms-statistics'); ?></th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <?php foreach ($form_order as $row_key) :
                                                $entry = $form_rows[$row_key];
                                                $label_title = $entry['title'];
                                                $checked = in_array($row_key, $selected_form_ids, true) ? 'checked' : '';
                                                $entry_count = $entry['count'];
                                                $ids = is_array($entry['ids']) ? $entry['ids'] : array();
                                                if ($display_mode === 'form_name_distinct') {
                                                    $id_label = implode(', ', $ids);
                                                } elseif ($display_mode === 'form_name') {
                                                    $parts = explode('||', $row_key, 2);
                                                    $id_label = $parts[0];
                                                } else {
                                                    $id_label = $row_key;
                                                }
                                            ?>
                                                <tr>
                                                    <th scope="row" class="check-column">
                                                        <label class="screen-reader-text"><?php printf(__('Auswählen %s', 'elementor-forms-statistics'), esc_html($label_title)); ?></label>
                                                        <input type="checkbox" name="selected_form_ids[]" value="<?php echo esc_attr($row_key); ?>" <?php echo $checked; ?>>
                                                    </th>
                                                    <td><?php echo esc_html($id_label); ?></td>
                                                    <td><?php echo esc_html($label_title); ?></td>
                                                    <td><?php echo esc_html(number_format_i18n($entry_count)); ?></td>
                                                </tr>
                                            <?php endforeach; ?>
                                        </tbody>
                                    </table>
                                </div>
                            <?php endif; ?>
                            <p class="description"><?php _e('Nur ausgewählte Formulare werden in Grafik und Tabellen berücksichtigt.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label for="excluded_emails"><?php _e('E-Mail Ignorelist', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <textarea name="excluded_emails" id="excluded_emails" rows="8" class="large-text code"><?php echo esc_textarea($stored_text); ?></textarea>
                            <p class="description"><?php _e('E-Mails in diesem Feld fließen nicht in die Statistik ein. Bitte je Zeile eine Adresse oder einen Teilstring eintragen; Wildcards wie * können verwendet werden.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label for="emphasis_threshold"><?php _e('Schwellwert für außergewöhnliche Veränderungen', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <input type="number" min="0" step="1" name="emphasis_threshold" id="emphasis_threshold" value="<?php echo esc_attr($emphasis_threshold); ?>">
                            <p class="description"><?php _e('Ab dieser Differenz zum Vorjahr werden Werte in der Tabelle fett hervorgehoben. Standard: 5.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label for="export_folder"><?php _e('Export-Ordner (Upload-Verzeichnis)', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <input type="text" name="export_folder" id="export_folder" value="<?php echo esc_attr($export_folder_slug); ?>" class="regular-text">
                            <p class="description"><?php _e('Ordnername unter wp-content/uploads für die gespeicherten Statistiken.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label><?php _e('Benutzerrollen', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <?php if (empty($menu_role_choices)) : ?>
                                <p><?php _e('Keine Benutzerrollen gefunden.', 'elementor-forms-statistics'); ?></p>
                            <?php else : ?>
                                <div class="mdp-role-matrix-wrap">
                                    <table class="widefat striped mdp-role-matrix">
                                        <thead>
                                            <tr>
                                                <th><?php _e('Menüpunkt', 'elementor-forms-statistics'); ?></th>
                                                <th><?php _e('Zugriff für Rollen', 'elementor-forms-statistics'); ?></th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <?php foreach ($menu_items as $menu_key => $menu_label) : ?>
                                                <tr>
                                                    <td><?php echo esc_html($menu_label); ?></td>
                                                    <td>
                                                        <?php foreach ($menu_role_choices as $role_slug => $role_label) : ?>
                                                            <label class="mdp-role-checkbox">
                                                                <input type="checkbox" name="menu_roles[<?php echo esc_attr($menu_key); ?>][]" value="<?php echo esc_attr($role_slug); ?>" <?php checked(in_array($role_slug, $menu_roles[$menu_key], true)); ?>>
                                                                <?php echo esc_html($role_label); ?>
                                                            </label>
                                                        <?php endforeach; ?>
                                                    </td>
                                                </tr>
                                            <?php endforeach; ?>
                                        </tbody>
                                    </table>
                                </div>
                            <?php endif; ?>
                            <p class="description"><?php _e('Lege fest, welche Benutzerrollen die einzelnen Menüeinträge von Elementor Forms Statistics sehen.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label><?php _e('Farben', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <div class="mdp-color-matrix">
                                <table class="mdp-color-table">
                                    <tbody>
                                    <?php foreach ($curve_color_slots as $index => $slot) : ?>
                                        <tr>
                                            <td class="mdp-color-label">
                                                <?php
                                                if ($index === 0) {
                                                    echo esc_html__('Aktuelles Jahr', 'elementor-forms-statistics');
                                                } else {
                                                    printf(esc_html__('%d Jahr', 'elementor-forms-statistics'), $index + 1);
                                                }
                                                ?>
                                            </td>
                                            <td class="mdp-color-input-cell">
                                                <input type="color" name="curve_color_<?php echo $index; ?>" value="<?php echo esc_attr($slot['color']); ?>" class="mdp-curve-color-input">
                                            </td>
                                            <td class="mdp-transparency-cell">
                                                <label class="mdp-transparency-label">
                                                    <?php _e('Transparenz', 'elementor-forms-statistics'); ?>:
                                                    <span class="mdp-transparency-value"><?php echo esc_html(round($slot['alpha'] * 100)); ?>%</span>
                                                </label>
                                                <input type="range" min="0" max="100" step="5" name="curve_alpha_<?php echo $index; ?>" value="<?php echo esc_attr(round($slot['alpha'] * 100)); ?>" class="mdp-transparency-slider">
                                            </td>
                                        </tr>
                                    <?php endforeach; ?>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label for="clean_on_uninstall"><?php _e('Verhalten Deinstallieren', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <label class="mdp-checkbox-inline">
                                <input type="checkbox" name="clean_on_uninstall" id="clean_on_uninstall" value="1" <?php checked($clean_on_uninstall, '1'); ?>>
                                <span class="mdp-clean-action-text">⚠️ <?php _e('Beim Deinstallieren dieses Plugins alle Daten bereinigen.', 'elementor-forms-statistics'); ?></span>
                            </label>
                            <p class="description mdp-warning-text"><?php _e('Wenn diese Checkbox aktiviert ist, wird bei der Deinstallation alles bereinigt: Datenbank-Tabellen, archivierte Statistiken und die statischen Dateien im Upload-Ordner werden gelöscht. Elementor-Submissions bleiben unberührt.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
                <div class="mdp-form-section">
                    <div class="mdp-form-field">
                        <label><?php _e('Elementor Submissions bereinigen', 'elementor-forms-statistics'); ?></label>
                        <div class="mdp-form-field-control">
                            <select name="submission_cleanup_interval">
                                <?php
                                $cleanup_options = array(
                                    'disabled' => __('Deaktiviert', 'elementor-forms-statistics'),
                                    '1h' => __('1 Stunde', 'elementor-forms-statistics'),
                                    '1d' => __('1 Tag', 'elementor-forms-statistics'),
                                    '1m' => __('1 Monat', 'elementor-forms-statistics'),
                                    '1y' => __('1 Jahr', 'elementor-forms-statistics'),
                                );
                                foreach ($cleanup_options as $key => $label) :
                                ?>
                                    <option value="<?php echo esc_attr($key); ?>" <?php selected($cleanup_interval, $key); ?>><?php echo esc_html($label); ?></option>
                                <?php endforeach; ?>
                            </select>
                            <p class="description mdp-cleanup-warning">⚠️ <?php _e('Nur wenn ein Intervall ausgewählt ist, werden alte Elementor Submissions im gewählten Rhythmus gelöscht, nachdem sie ins Archiv übernommen wurden.', 'elementor-forms-statistics'); ?></p>
                        </div>
                    </div>
                </div>
            </div>
            <?php submit_button(__('Speichern', 'elementor-forms-statistics')); ?>
        </form>
        <script>
        (function($){
            function mdpUpdateTransparencyValue(input) {
                var $input = $(input);
                var $value = $input.closest('.mdp-transparency-cell').find('.mdp-transparency-value');
                if ($value.length) {
                    $value.text($input.val() + '%');
                }
            }
            $(document).on('input', '.mdp-transparency-slider', function() {
                mdpUpdateTransparencyValue(this);
            });
            $(document).ready(function() {
                $('.mdp-transparency-slider').each(function() {
                    mdpUpdateTransparencyValue(this);
                });
            });
        })(jQuery);
        </script>
    </div>
    <?php
}

function mdp_archive_page_callback() {
    if (!mdp_user_can_access_menu('statistiken-archiv')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }

    $status = isset($_GET['mdp_archive_status']) ? sanitize_text_field($_GET['mdp_archive_status']) : '';
    $table_exists = mdp_archive_table_exists();
    $initialized = (bool) get_option('mdp_efs_archive_initialized');
    $last_run_raw = get_option('mdp_efs_archive_last_run', '');
    $last_sync_timestamp = (int) get_option('mdp_efs_archive_last_sync', 0);
    $last_run_display = $last_run_raw ? date_i18n(get_option('date_format') . ' ' . get_option('time_format'), strtotime($last_run_raw)) : __('Noch kein Import durchgeführt.', 'elementor-forms-statistics');
    $last_sync_display = $last_sync_timestamp ? sprintf(__('%s (vor %s)', 'elementor-forms-statistics'), date_i18n(get_option('date_format') . ' ' . get_option('time_format'), $last_sync_timestamp), human_time_diff($last_sync_timestamp, current_time('timestamp'))) : __('Noch nie', 'elementor-forms-statistics');

    $row_count = 0;
    $distinct_forms = 0;
    $total_requests = 0;
    if ($table_exists) {
        global $wpdb;
        $table = mdp_get_archive_table_name();
        $row_count = (int) $wpdb->get_var("SELECT COUNT(*) FROM {$table}");
        $distinct_forms = (int) $wpdb->get_var("SELECT COUNT(DISTINCT form_id) FROM {$table}");
        $total_requests = (int) $wpdb->get_var("SELECT SUM(total) FROM {$table}");
    }

    $button_label = $initialized
        ? __('Neue Einträge sichern', 'elementor-forms-statistics')
        : __('Archiv initialisieren', 'elementor-forms-statistics');
    ?>
    <div class="wrap mdp-stats-root">
        <h1><?php _e('Archivierte Statistiken', 'elementor-forms-statistics'); ?></h1>
        <?php if ($status === 'initialized') : ?>
            <div class="notice notice-success"><p><?php _e('Das Archiv wurde vollständig aufgebaut.', 'elementor-forms-statistics'); ?></p></div>
        <?php elseif ($status === 'synced') : ?>
            <div class="notice notice-success"><p><?php _e('Neue Einträge wurden ins Archiv übernommen.', 'elementor-forms-statistics'); ?></p></div>
        <?php endif; ?>
        <p><?php _e('Das Archiv speichert aggregierte und anonymisierte Formulardaten. So bleiben Statistiken erhalten, auch wenn Elementor-Einträge gelöscht oder aufgeräumt werden.', 'elementor-forms-statistics'); ?></p>
        <table class="form-table" role="presentation">
            <tr>
                <th scope="row"><?php _e('Status', 'elementor-forms-statistics'); ?></th>
                <td>
                    <?php if ($initialized && $table_exists) : ?>
                        <span class="dashicons dashicons-yes" style="color: #46b450;"></span>
                        <?php _e('Aktiv', 'elementor-forms-statistics'); ?>
                    <?php else : ?>
                        <span class="dashicons dashicons-warning" style="color: #d63638;"></span>
                        <?php _e('Noch nicht initialisiert', 'elementor-forms-statistics'); ?>
                    <?php endif; ?>
                </td>
            </tr>
            <tr>
                <th scope="row"><?php _e('Letzter vollständiger Import', 'elementor-forms-statistics'); ?></th>
                <td><?php echo esc_html($last_run_display); ?></td>
            </tr>
            <tr>
                <th scope="row"><?php _e('Letzte Synchronisierung', 'elementor-forms-statistics'); ?></th>
                <td><?php echo esc_html($last_sync_display); ?></td>
            </tr>
            <tr>
                <th scope="row"><?php _e('Gespeicherte Monate', 'elementor-forms-statistics'); ?></th>
                <td><?php echo esc_html(number_format_i18n($row_count)); ?></td>
            </tr>
            <tr>
                <th scope="row"><?php _e('Anzahl unterschiedlicher Formulare', 'elementor-forms-statistics'); ?></th>
                <td><?php echo esc_html(number_format_i18n($distinct_forms)); ?></td>
            </tr>
            <tr>
                <th scope="row"><?php _e('Summe aller Anfragen', 'elementor-forms-statistics'); ?></th>
                <td><?php echo esc_html(number_format_i18n($total_requests)); ?></td>
            </tr>
        </table>
        <p><?php _e('Nach der Initialisierung synchronisiert das Plugin automatisch neue Einträge in kurzen Abständen. Bei Bedarf kann der Prozess hier jederzeit erneut gestartet werden.', 'elementor-forms-statistics'); ?></p>
        <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
            <?php wp_nonce_field('mdp_archive_import', 'mdp_archive_nonce'); ?>
            <input type="hidden" name="action" value="mdp_run_archive_import">
            <?php if ($initialized) : ?>
                <label>
                    <input type="checkbox" name="mdp_archive_reset" value="1">
                    <?php _e('Archiv vollständig neu aufbauen (überschreibt bestehende Daten mit aktuellem Elementor-Stand)', 'elementor-forms-statistics'); ?>
                </label>
                <p class="description"><?php _e('Nur aktivieren, wenn das Archiv beschädigt ist. Achtung: Gelöschte Elementor-Einträge gehen dabei verloren.', 'elementor-forms-statistics'); ?></p>
            <?php endif; ?>
            <?php submit_button($button_label); ?>
        </form>
    </div>
    <?php
}

function mdp_export_page_callback() {
    if (!mdp_user_can_access_export_menu()) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }
    $forms_indexed = mdp_get_forms_indexed();
    $form_ids = array_keys($forms_indexed);
    sort($form_ids, SORT_STRING);
    $selected_form_id = isset($_GET['form_id']) ? sanitize_text_field(wp_unslash($_GET['form_id'])) : '';
    if ($selected_form_id === '') {
        $selected_form_id = mdp_get_last_export_form_id(get_current_user_id());
    }
    if ($selected_form_id === '' && !empty($form_ids)) {
        $selected_form_id = $form_ids[0];
    }
    if ($selected_form_id !== '' && !isset($forms_indexed[$selected_form_id]) && !empty($form_ids)) {
        $selected_form_id = $form_ids[0];
    }
    $saved_notice = !empty($_GET['mdp_export_saved']);
    $script_path = plugin_dir_path(__FILE__) . 'assets/js/export.js';
    $script_version = file_exists($script_path) ? (string) filemtime($script_path) : '1.0.0';
    $export_config = array(
        'ajaxUrl' => admin_url('admin-ajax.php'),
        'nonce' => wp_create_nonce('mdp_export_fields'),
        'formId' => $selected_form_id,
        'dateFormats' => array_map(function($value, $label) {
            return array('value' => $value, 'label' => $label);
        }, array_keys(mdp_get_export_date_format_choices()), mdp_get_export_date_format_choices()),
        'i18n' => array(
            'saved' => __('Gespeichert', 'elementor-forms-statistics'),
            'no_fields' => __('Keine Felder gefunden.', 'elementor-forms-statistics'),
            'no_columns' => __('Keine Spalten ausgewählt.', 'elementor-forms-statistics'),
            'no_columns_body' => __('Bitte mindestens eine Export-Spalte aktivieren.', 'elementor-forms-statistics'),
            'no_entries' => __('Keine Einträge gefunden.', 'elementor-forms-statistics'),
            'delete_field' => __('Feld löschen', 'elementor-forms-statistics'),
            'delete_rule' => __('Regel löschen', 'elementor-forms-statistics'),
            'delete_formula' => __('Formel löschen', 'elementor-forms-statistics'),
        ),
    );
    wp_enqueue_script('jquery');
    wp_enqueue_script('jquery-ui-sortable');
    wp_print_scripts(array('jquery', 'jquery-ui-sortable'));
    echo '<script>window.mdpExportConfig = ' . wp_json_encode($export_config) . ';</script>';
    echo '<script src="' . esc_url(plugin_dir_url(__FILE__) . 'assets/js/export.js?ver=' . rawurlencode($script_version)) . '"></script>';
    $inline_url = wp_nonce_url(
        add_query_arg(
            array(
                'action' => 'mdp_export_stats_html',
                'inline' => 1,
            ),
            admin_url('admin-post.php')
        ),
        'mdp_export_html_inline',
        'mdp_inline_nonce'
    );
    ?>
    <div class="wrap mdp-stats-root">
        <h1><?php _e('Formular-Daten Export', 'elementor-forms-statistics'); ?></h1>
        <?php if ($saved_notice) : ?>
            <div class="notice notice-success"><p><?php _e('Export-Einstellungen gespeichert.', 'elementor-forms-statistics'); ?></p></div>
        <?php endif; ?>
        <?php if (empty($forms_indexed)) : ?>
            <p><?php _e('Keine Formulare gefunden.', 'elementor-forms-statistics'); ?></p>
        <?php else : ?>
            <div class="mdp-export-tabs">
                <button type="button" class="mdp-export-tab is-active" data-tab="mdp-export-tab-form">
                    <?php _e('Formular auswählen', 'elementor-forms-statistics'); ?>
                </button>
                <button type="button" class="mdp-export-tab" data-tab="mdp-export-tab-columns">
                    <?php _e('Export-Spalten & Vorschau', 'elementor-forms-statistics'); ?>
                </button>
                <button type="button" class="mdp-export-tab" data-tab="mdp-export-tab-rules">
                    <?php _e('Regeln & Formeln', 'elementor-forms-statistics'); ?>
                </button>
            </div>
            <div id="mdp-export-tab-form" class="mdp-export-tab-panel is-active">
                <div class="mdp-export-col mdp-export-col-left">
                    <h2><?php _e('Formular auswählen', 'elementor-forms-statistics'); ?></h2>
                    <label for="mdp_export_form_id"><?php _e('Formular', 'elementor-forms-statistics'); ?></label>
                    <select id="mdp_export_form_id">
                        <?php foreach ($form_ids as $form_id) :
                            $form_entry = $forms_indexed[$form_id];
                            $title = mdp_format_form_title($form_id, $forms_indexed, mdp_resolve_form_title($form_id, $form_entry));
                            $label = $title !== '' ? $title : $form_id;
                        ?>
                            <option value="<?php echo esc_attr($form_id); ?>" <?php selected($selected_form_id, $form_id); ?>>
                                <?php echo esc_html($label . ' (' . $form_id . ')'); ?>
                            </option>
                        <?php endforeach; ?>
                    </select>
                    <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="mdp-export-form">
                        <?php wp_nonce_field('mdp_export_csv', 'mdp_export_csv_nonce'); ?>
                        <input type="hidden" name="action" value="mdp_export_csv">
                        <input type="hidden" name="form_id" id="mdp_export_form_id_csv" value="<?php echo esc_attr($selected_form_id); ?>">
                        <?php
                        // Single export action (CSV only).
                        submit_button(__('Exportieren', 'elementor-forms-statistics'), 'primary');
                        ?>
                    </form>
                </div>
            </div>
            <div id="mdp-export-tab-columns" class="mdp-export-tab-panel">
                <div class="mdp-export-col mdp-export-col-right">
                    <div class="mdp-export-right-grid">
                        <div class="mdp-export-fields-panel">
                            <h2><?php _e('Export-Spalten', 'elementor-forms-statistics'); ?></h2>
                            <p><?php _e('Felder per Drag & Drop sortieren und Spaltennamen anpassen.', 'elementor-forms-statistics'); ?></p>
                            <div class="mdp-export-add-field">
                                <label for="mdp_export_new_field"><?php _e('Eigenes Feld hinzufügen', 'elementor-forms-statistics'); ?></label>
                                <div class="mdp-export-add-field-row">
                                    <input type="text" id="mdp_export_new_field" class="regular-text" placeholder="<?php echo esc_attr__('Feldname', 'elementor-forms-statistics'); ?>">
                                    <button type="button" class="button" id="mdp_export_add_field"><?php _e('Hinzufügen', 'elementor-forms-statistics'); ?></button>
                                </div>
                            </div>
                            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" id="mdp-export-fields-form">
                                <?php wp_nonce_field('mdp_save_export_fields', 'mdp_export_fields_nonce'); ?>
                                <input type="hidden" name="action" value="mdp_save_export_fields">
                                <input type="hidden" name="form_id" id="mdp_export_form_id_save" value="<?php echo esc_attr($selected_form_id); ?>">
                                <input type="hidden" name="fields_payload" id="mdp_export_fields_payload" value="">
                                <table class="mdp-export-fields-table">
                                    <thead>
                                        <tr>
                                    <th class="mdp-export-col-handle"></th>
                                    <th><?php _e('Feld', 'elementor-forms-statistics'); ?></th>
                                    <th><?php _e('Spaltenname', 'elementor-forms-statistics'); ?></th>
                                    <th><?php _e('Datumformat', 'elementor-forms-statistics'); ?></th>
                                    <th class="mdp-export-col-include"><?php _e('Export', 'elementor-forms-statistics'); ?></th>
                                </tr>
                            </thead>
                                    <tbody id="mdp-export-fields-list"></tbody>
                                </table>
                                <?php submit_button(__('Reihenfolge speichern', 'elementor-forms-statistics'), 'secondary'); ?>
                            </form>
                        </div>
                        <div class="mdp-export-preview-panel">
                            <h3><?php _e('Vorschau', 'elementor-forms-statistics'); ?></h3>
                        <div class="mdp-export-preview">
                                <table class="mdp-export-preview-table">
                                    <?php
                                    $preview_data = mdp_get_export_preview_data($selected_form_id);
                                    $preview_headers = isset($preview_data['headers']) ? $preview_data['headers'] : [];
                                    $preview_rows = isset($preview_data['rows']) ? $preview_data['rows'] : [];
                                    ?>
                                    <thead>
                                        <tr>
                                            <?php if (empty($preview_headers)) : ?>
                                                <th><?php _e('Keine Spalten ausgewählt.', 'elementor-forms-statistics'); ?></th>
                                            <?php else : ?>
                                                <?php foreach ($preview_headers as $header_label) : ?>
                                                    <th><?php echo esc_html($header_label); ?></th>
                                                <?php endforeach; ?>
                                            <?php endif; ?>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <?php if (empty($preview_headers)) : ?>
                                            <tr>
                                                <td><?php _e('Bitte mindestens eine Export-Spalte aktivieren.', 'elementor-forms-statistics'); ?></td>
                                            </tr>
                                        <?php elseif (empty($preview_rows)) : ?>
                                            <tr>
                                                <td colspan="<?php echo (int) count($preview_headers); ?>"><?php _e('Keine Einträge gefunden.', 'elementor-forms-statistics'); ?></td>
                                            </tr>
                                        <?php else : ?>
                                            <?php foreach ($preview_rows as $row) : ?>
                                                <tr>
                                                    <?php foreach ($row as $cell) : ?>
                                                        <td><?php echo esc_html($cell); ?></td>
                                                    <?php endforeach; ?>
                                                </tr>
                                            <?php endforeach; ?>
                                        <?php endif; ?>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div id="mdp-export-tab-rules" class="mdp-export-tab-panel">
                <div class="mdp-export-col mdp-export-col-right">
                    <div class="mdp-export-right-grid">
                        <div class="mdp-export-fields-panel">
                            <div class="mdp-export-rules">
                                <h3><?php _e('Suchen & Ersetzen', 'elementor-forms-statistics'); ?></h3>
                                <p><?php _e('Regeln werden beim Export und in der Vorschau angewendet.', 'elementor-forms-statistics'); ?></p>
                                <table class="mdp-export-rules-table">
                                    <thead>
                                        <tr>
                                            <th><?php _e('Suchen', 'elementor-forms-statistics'); ?></th>
                                            <th><?php _e('Ersetzen', 'elementor-forms-statistics'); ?></th>
                                            <th></th>
                                        </tr>
                                    </thead>
                                    <tbody id="mdp-export-rules-body"></tbody>
                                </table>
                                <button type="button" class="button" id="mdp_export_add_rule"><?php _e('Regel hinzufügen', 'elementor-forms-statistics'); ?></button>
                            </div>
                            <div class="mdp-export-formulas">
                                <h3><?php _e('Formel erstellen', 'elementor-forms-statistics'); ?></h3>
                                <div class="mdp-export-formula-add">
                                    <input type="text" id="mdp_export_new_formula_label" class="regular-text" placeholder="<?php echo esc_attr__('Feldname', 'elementor-forms-statistics'); ?>">
                                    <button type="button" class="button" id="mdp_export_add_formula"><?php _e('Formel erstellen', 'elementor-forms-statistics'); ?></button>
                                </div>
                                <div class="mdp-formula-tags">
                                    <span class="mdp-formula-tags-label"><?php _e('Felder', 'elementor-forms-statistics'); ?></span>
                                    <div id="mdp-formula-tags-list" class="mdp-formula-tags-list"></div>
                                </div>
                                <table class="mdp-export-formulas-table">
                                    <thead>
                                        <tr>
                                            <th><?php _e('Feldname', 'elementor-forms-statistics'); ?></th>
                                            <th><?php _e('Formel', 'elementor-forms-statistics'); ?></th>
                                            <th></th>
                                        </tr>
                                    </thead>
                                    <tbody id="mdp-export-formulas-body"></tbody>
                                </table>
                            </div>
                        </div>
                        <div class="mdp-export-preview-panel">
                            <h3><?php _e('Vorschau', 'elementor-forms-statistics'); ?></h3>
                            <div class="mdp-export-preview">
                                <table class="mdp-export-preview-table">
                                    <thead></thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        <?php endif; ?>
    </div>
    <?php if (!empty($forms_indexed)) : ?>
    <?php endif; ?>
    <?php
}

function mdp_email_settings_page_callback() {
    if (!mdp_user_can_access_menu('statistiken-emailversand')) {
        wp_die(__('Sie haben keine Berechtigung, diese Seite zu sehen.', 'elementor-forms-statistics'));
    }

    $message = '';
    if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['mdp_email_settings_nonce']) && wp_verify_nonce($_POST['mdp_email_settings_nonce'], 'mdp_save_email_settings')) {
        $email_interval = isset($_POST['email_interval']) ? sanitize_text_field($_POST['email_interval']) : 'monthly';
        $allowed_intervals = array('disabled', 'daily', 'weekly', 'monthly');
        if (!in_array($email_interval, $allowed_intervals, true)) {
            $email_interval = 'monthly';
        }

        $email_recipients_raw = isset($_POST['email_recipients']) ? wp_unslash($_POST['email_recipients']) : '';
        $email_recipient_lines = preg_split('/[\r\n,]+/', $email_recipients_raw);
        $clean_recipients = [];
        foreach ($email_recipient_lines as $recipient_line) {
            $recipient_line = trim($recipient_line);
            if ($recipient_line === '') {
                continue;
            }
            $recipient_email = sanitize_email($recipient_line);
            if ($recipient_email) {
                $clean_recipients[] = strtolower($recipient_email);
            }
        }

        $email_message_input = isset($_POST['email_message']) ? wp_unslash($_POST['email_message']) : '';
        $email_message_input = is_string($email_message_input) ? trim($email_message_input) : '';
        $email_message = $email_message_input !== '' ? wp_kses_post($email_message_input) : mdp_get_default_email_message();
        $email_subject = isset($_POST['email_subject']) ? sanitize_text_field(wp_unslash($_POST['email_subject'])) : '';
        if ($email_subject === '') {
            $email_subject = '📈 ' . __('Anfragen Statistik – %s', 'elementor-forms-statistics');
        }

        $export_link_label = isset($_POST['export_link_label']) ? sanitize_text_field(wp_unslash($_POST['export_link_label'])) : '';
        $email_send_time = isset($_POST['email_send_time']) ? sanitize_text_field($_POST['email_send_time']) : '08:00';
        if (!preg_match('/^([01]\d|2[0-3]):([0-5]\d)$/', $email_send_time)) {
            $email_send_time = '08:00';
        }
        $email_day_of_month = isset($_POST['email_day_of_month']) ? (int) $_POST['email_day_of_month'] : 1;
        if ($email_day_of_month < 1) {
            $email_day_of_month = 1;
        } elseif ($email_day_of_month > 31) {
            $email_day_of_month = 31;
        }
        $email_weekday = isset($_POST['email_weekday']) ? strtolower(sanitize_text_field($_POST['email_weekday'])) : 'monday';
        $weekday_choices = array_keys(mdp_get_email_weekday_choices());
        if (!in_array($email_weekday, $weekday_choices, true)) {
            $email_weekday = 'monday';
        }

        $include_attachment = isset($_POST['include_attachment']) && $_POST['include_attachment'] === '1' ? '1' : '0';

        update_option('mdp_efs_email_interval', $email_interval);
        update_option('mdp_efs_email_recipients', implode("\n", $clean_recipients));
        update_option('mdp_efs_email_message', $email_message);
        update_option('mdp_efs_email_subject', $email_subject);
        update_option('mdp_efs_email_time', $email_send_time);
        update_option('mdp_efs_email_day_of_month', $email_day_of_month);
        update_option('mdp_efs_email_weekday', $email_weekday);
        update_option('mdp_efs_export_link_label', $export_link_label);
        update_option('mdp_efs_include_attachment', $include_attachment);
        mdp_reset_stats_email_schedule();
        $message = __('Einstellungen gespeichert.', 'elementor-forms-statistics');
    }

    $email_interval = mdp_get_schedule_interval();
    $email_recipients_text = mdp_get_email_recipients_text();
    $email_message = mdp_get_email_message();
    $email_subject_template = mdp_get_email_subject_template();
    $email_send_time = mdp_get_email_send_time();
    $email_day_of_month = mdp_get_email_day_of_month();
    $email_weekday = mdp_get_email_weekday();
    $manual_status = isset($_GET['mdp_send_status']) ? sanitize_text_field($_GET['mdp_send_status']) : '';
    $weekday_choices = mdp_get_email_weekday_choices();
    $link_label = mdp_get_export_link_label();
    $include_attachment = mdp_should_include_stats_attachment();
    $show_time_row = $email_interval !== 'disabled';
    $show_weekly_row = ($email_interval === 'weekly');
    $show_monthly_row = ($email_interval === 'monthly');
    ?>
    <div class="wrap mdp-stats-root">
        <h1><?php _e('E-Mail Versand', 'elementor-forms-statistics'); ?></h1>
        <?php if ($message) : ?>
            <div class="notice notice-success"><p><?php echo esc_html($message); ?></p></div>
        <?php endif; ?>
        <?php if ($manual_status === 'success') : ?>
            <div class="notice notice-success"><p><?php _e('Die Statistik wurde erfolgreich versendet.', 'elementor-forms-statistics'); ?></p></div>
        <?php elseif ($manual_status === 'error') : ?>
            <div class="notice notice-error"><p><?php _e('Die Statistik konnte nicht gesendet werden. Bitte Empfänger prüfen.', 'elementor-forms-statistics'); ?></p></div>
        <?php endif; ?>
        <div class="mdp-email-tabs" role="tablist">
            <button type="button" class="mdp-email-tab is-active" data-panel="planning"><?php _e('Planung', 'elementor-forms-statistics'); ?></button>
            <button type="button" class="mdp-email-tab" data-panel="email"><?php _e('E-Mail', 'elementor-forms-statistics'); ?></button>
            <button type="button" class="mdp-email-tab" data-panel="send"><?php _e('Senden & Download', 'elementor-forms-statistics'); ?></button>
        </div>
        <form method="post">
            <?php wp_nonce_field('mdp_save_email_settings', 'mdp_email_settings_nonce'); ?>
            <div class="mdp-form-wrapper mdp-email-form-wrapper">
                <div class="mdp-email-tab-panels">
                    <div class="mdp-email-tab-panel is-active" data-panel="planning">
                        <div class="mdp-form-section">
                            <label for="email_interval"><?php _e('Automatischer Versand', 'elementor-forms-statistics'); ?></label>
                            <div class="mdp-form-field-control">
                                <select name="email_interval" id="email_interval">
                                    <option value="disabled" <?php selected($email_interval, 'disabled'); ?>><?php _e('Deaktiviert', 'elementor-forms-statistics'); ?></option>
                                    <option value="daily" <?php selected($email_interval, 'daily'); ?>><?php _e('Täglich', 'elementor-forms-statistics'); ?></option>
                                    <option value="weekly" <?php selected($email_interval, 'weekly'); ?>><?php _e('Wöchentlich', 'elementor-forms-statistics'); ?></option>
                                    <option value="monthly" <?php selected($email_interval, 'monthly'); ?>><?php _e('Monatlich', 'elementor-forms-statistics'); ?></option>
                                </select>
                                <p class="description"><?php _e('Legt fest, wie oft die Statistik per E-Mail verschickt wird.', 'elementor-forms-statistics'); ?></p>
                            </div>
                        </div>
                        <div class="mdp-form-section mdp-schedule-row"<?php echo $show_time_row || $show_weekly_row ? '' : ' style="display:none;"'; ?>>
                            <div class="mdp-form-column mdp-field-time"<?php echo $show_time_row ? '' : ' style="display:none;"'; ?>>
                                <div class="mdp-form-field">
                                    <label for="email_send_time"><?php _e('Uhrzeit für den Versand', 'elementor-forms-statistics'); ?></label>
                                    <div class="mdp-form-field-control">
                                        <input type="time" name="email_send_time" id="email_send_time" value="<?php echo esc_attr($email_send_time); ?>">
                                        <p class="description"><?php _e('Die Statistik wird zur angegebenen Uhrzeit versendet.', 'elementor-forms-statistics'); ?></p>
                                    </div>
                                </div>
                            </div>
                            <div class="mdp-form-column mdp-field-weekly"<?php echo $show_weekly_row ? '' : ' style="display:none;"'; ?>>
                                <div class="mdp-form-field">
                                    <label for="email_weekday"><?php _e('Wochentag für den Versand', 'elementor-forms-statistics'); ?></label>
                                    <div class="mdp-form-field-control">
                                        <select name="email_weekday" id="email_weekday">
                                            <?php foreach ($weekday_choices as $weekday_slug => $weekday_label) : ?>
                                                <option value="<?php echo esc_attr($weekday_slug); ?>" <?php selected($email_weekday, $weekday_slug); ?>><?php echo esc_html($weekday_label); ?></option>
                                            <?php endforeach; ?>
                                        </select>
                                        <p class="description"><?php _e('Nur sichtbar bei wöchentlichem Versand.', 'elementor-forms-statistics'); ?></p>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="mdp-form-section mdp-field-monthly"<?php echo $show_monthly_row ? '' : ' style="display:none;"'; ?>>
                            <div class="mdp-form-field">
                                <label for="email_day_of_month"><?php _e('Tag im Monat', 'elementor-forms-statistics'); ?></label>
                                <div class="mdp-form-field-control">
                                    <input type="number" min="1" max="31" name="email_day_of_month" id="email_day_of_month" value="<?php echo esc_attr($email_day_of_month); ?>">
                                    <p class="description"><?php _e('Bei unterschiedlichen Monatslängen wird automatisch der letzte verfügbare Tag verwendet.', 'elementor-forms-statistics'); ?></p>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="mdp-email-tab-panel" data-panel="email">
                        <div class="mdp-form-section">
                            <label for="email_recipients"><?php _e('Empfänger der Statistik', 'elementor-forms-statistics'); ?></label>
                            <div class="mdp-form-field-control">
                                <textarea name="email_recipients" id="email_recipients" rows="4" class="large-text code"><?php echo esc_textarea($email_recipients_text); ?></textarea>
                                <p class="description"><?php _e('Eine oder mehrere E-Mail-Adressen, getrennt durch Zeilenumbrüche oder Kommas.', 'elementor-forms-statistics'); ?></p>
                            </div>
                        </div>
                        <div class="mdp-form-section mdp-form-section--no-divider">
                            <label for="email_subject"><?php _e('Betreffzeile', 'elementor-forms-statistics'); ?></label>
                            <div class="mdp-form-field-control">
                                <input type="text" name="email_subject" id="email_subject" value="<?php echo esc_attr($email_subject_template); ?>" class="regular-text">
                            </div>
                        </div>
                        <div class="mdp-form-section mdp-form-section--no-divider">
                            <label for="email_message"><?php _e('E-Mail Text', 'elementor-forms-statistics'); ?></label>
                            <div class="mdp-form-field-control">
                                <textarea name="email_message" id="email_message" rows="6" class="large-text code"><?php echo esc_textarea($email_message); ?></textarea>
                                <p class="description"><?php _e('Der komplette Text für den E-Mail-Body.', 'elementor-forms-statistics'); ?></p>
                                <p class="description"><?php printf(
                                    __('%s wird durch den Link zur Statistik ersetzt. Fehlt der Platzhalter, hängt das System den Link am Ende an.', 'elementor-forms-statistics'),
                                    '<code>' . esc_html(mdp_get_stats_link_placeholder()) . '</code>'
                                ); ?></p>
                                <p class="description"><?php _e('%s = Website URL', 'elementor-forms-statistics'); ?></p>
                            </div>
                        </div>
                        <div class="mdp-form-section mdp-form-section--no-divider">
                            <label for="export_link_label"><?php _e('Text Link', 'elementor-forms-statistics'); ?></label>
                            <div class="mdp-form-field-control">
                                <input type="text" name="export_link_label" id="export_link_label" value="<?php echo esc_attr($link_label); ?>" class="regular-text">
                                <p class="description"><?php _e('Dieser Text ersetzt „Passwortfreien Statistik-Zugang“ im Link, falls du ihn anders benennen möchtest.', 'elementor-forms-statistics'); ?></p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <?php submit_button(__('Speichern', 'elementor-forms-statistics')); ?>
        </form>
        <div class="mdp-email-tab-panel" data-panel="send">
            <div class="mdp-form-wrapper mdp-email-form-wrapper">
                <div class="mdp-form-section mdp-form-section--no-divider">
                    <h3><?php _e('Statistik sofort versenden', 'elementor-forms-statistics'); ?></h3>
                    <p><?php _e('Verschickt die aktuelle Statistik einmalig an die oben definierten Empfänger.', 'elementor-forms-statistics'); ?></p>
                    <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="mdp-send-now-form">
                        <?php wp_nonce_field('mdp_send_stats_now', 'mdp_send_now_nonce'); ?>
                        <input type="hidden" name="action" value="mdp_send_stats_now">
                        <?php submit_button(__('Statistik jetzt senden', 'elementor-forms-statistics'), 'secondary', 'mdp_send_now'); ?>
                    </form>
                </div>
                <div class="mdp-form-section mdp-form-section--no-divider">
                    <h3><?php _e('Statistik als HTML exportieren', 'elementor-forms-statistics'); ?></h3>
                    <?php if (mdp_user_can_access_menu('statistiken-emailversand')) : ?>
                        <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" class="mdp-export-form">
                            <?php wp_nonce_field('mdp_export_html', 'mdp_export_nonce'); ?>
                            <input type="hidden" name="action" value="mdp_export_stats_html">
                            <?php submit_button(__('Als HTML exportieren', 'elementor-forms-statistics'), 'secondary'); ?>
                        </form>
                    <?php else : ?>
                        <p><?php _e('Der HTML-Export ist für deine Benutzerrolle nicht freigeschaltet.', 'elementor-forms-statistics'); ?></p>
                    <?php endif; ?>
                </div>
            </div>
        </div>
        <script>
        (function($){
            function mdpToggleScheduleFields() {
                var interval = $('#email_interval').val();
                $('.mdp-field-time').toggle(interval !== 'disabled');
                $('.mdp-field-weekly').toggle(interval === 'weekly');
                $('.mdp-field-monthly').toggle(interval === 'monthly');
            }
            function mdpSwitchEmailTab(panel) {
                var $tabs = $('.mdp-email-tab');
                var $panels = $('.mdp-email-tab-panel');
                $tabs.removeClass('is-active');
                $panels.removeClass('is-active');
                $tabs.filter('[data-panel="' + panel + '"]').addClass('is-active');
                $panels.filter('[data-panel="' + panel + '"]').addClass('is-active');
            }
            $(document).on('click', '.mdp-email-tab', function() {
                mdpSwitchEmailTab($(this).data('panel'));
            });
            $(document).ready(function() {
                $('.mdp-email-tab').first().trigger('click');
                $('.mdp-transparency-slider').each(function() {
                    var $input = $(this);
                    var $value = $input.closest('.mdp-transparency-cell').find('.mdp-transparency-value');
                    if ($value.length) {
                        $value.text($input.val() + '%');
                    }
                });
                mdpToggleScheduleFields();
            });
            $(document).on('change', '#email_interval', mdpToggleScheduleFields);
            $(document).on('input', '#email_interval', mdpToggleScheduleFields);
        })(jQuery);
        </script>
    </div>
    <?php
}

// Enable translation functions
add_action('plugins_loaded', function() {
    load_plugin_textdomain('elementor-forms-statistics', false, dirname(plugin_basename(__FILE__)) . '/languages/');
});

/****** Show Elmentor Submissions to editors **/
if (!class_exists('ElementorFormSubmissionsAccess'))
{

    class ElementorFormSubmissionsAccess
    {
        /** * See if this user is just an editor (if they have edit_posts but not manage_options).
         * If they have manage_options, they can see the Form Submissions page anyway.
         * @return boolean
         */
        static function isJustEditor()
        {
            return current_user_can('edit_posts') && !current_user_can('manage_options');
        }

        /**
         * This is called around line 849 of wp-includes/rest-api/class-wp-rest-server.php by the ajax request which loads the data
         * into the form submissions view for Elementor (see the add_menu_page below). The ajax request checks the user has
         * the manage_options permission in modules/forms/submissions/data/controller.php within the handler's permission_callback.
         * This overrides that, and also for the call to modules/forms/submissions/data/forms-controller.php (which fills the
         * Forms dropdown on the submissions page). By changing the $route check below, you could open up more pages to editors.
         * @param array [endpoints=>hanlders]
         * @return array [endpoints=>hanlders]
         */
        static function filterRestEndpoints($endpoints)
        {
            if (self::isJustEditor()) 
            {
                error_reporting(0); // there are a couple of PHP notices which prevent the Ajax JSON data from loading
                foreach($endpoints as $route=>$handlers) //for each endpoint
                    if (strpos($route, '/elementor/v1/form') === 0) //it is one of the elementor endpoints forms, form-submissions or form-submissions/export
                        foreach($handlers as $num=>$handler) //loop through the handlers
                            if (is_array ($handler) && isset ($handler['permission_callback'])) //if this handler has a permission_callback
                                $endpoints[$route][$num]['permission_callback'] = function($request){return true;}; //handler always returns true to grant permission
            }
            return $endpoints;
        }

        /**
         * Add the submissions page to the admin menu on the left for editors only, as administrators
         * can already see it.
         */
        static function addOptionsPage()
        {
            if (!self::isJustEditor()) return;
add_menu_page('Anfragen', 'Anfragen', 'edit_posts', 'e-form-submissions', function(){echo '<div id="e-form-submissions"></div>';}, 'dashicons-list-view', 3);


        }

        /**
         * Hook up the filter and action. I can't check if they are an editor here as the wp_user_can function
         * is not available yet.
         */
        static function hookIntoWordpress()
        {
            add_filter ('rest_endpoints', array('ElementorFormSubmissionsAccess', 'filterRestEndpoints'), 1, 3);
            add_action ('admin_menu', array('ElementorFormSubmissionsAccess', 'addOptionsPage'));
        }
    }

    ElementorFormSubmissionsAccess::hookIntoWordpress();
} //a wrapper to see if the class already exists or not



/* json */
// REST-API-Route registrieren
add_action('rest_api_init', function () {
    register_rest_route('mdp/v1', '/submissions', [
        'methods' => 'GET',
        'callback' => 'mdp_get_submission_data',
        'permission_callback' => function () {
            return current_user_can('edit_posts'); // Zugriffsberechtigungen für Autoren und höher
        }
    ]);
});

function mdp_get_submission_data($request) {
    global $wpdb;
    $table_name = $wpdb->prefix . 'e_submissions';
    $selected_form_id = sanitize_text_field($request->get_param('form_id'));
    if (!$selected_form_id) {
        $selected_form_id = sanitize_text_field($request->get_param('element_id'));
    }
    $selected_referer_title = sanitize_text_field($request->get_param('referer_title')); // Legacy fallback

    // Basisabfrage
    $sql = "SELECT YEAR(created_at_gmt) AS jahr, MONTH(created_at_gmt) AS monat, COUNT(*) AS anzahl_anfragen 
            FROM " . $table_name . " 
            WHERE `status` NOT LIKE '%trash%'";

    // Formular-Filter anwenden
    if ($selected_form_id) {
        $sql .= $wpdb->prepare(" AND `element_id` = %s", $selected_form_id);
    } elseif ($selected_referer_title) { // Legacy fallback to avoid breaking existing clients
        $sql .= $wpdb->prepare(" AND `referer_title` = %s", $selected_referer_title);
    }

    $sql .= mdp_get_email_exclusion_clause($table_name);

    $sql .= " GROUP BY jahr, monat ORDER BY jahr, monat";

    $results = $wpdb->get_results($sql);

    // Datenstruktur für JSON-Ausgabe vorbereiten
    $data = [];
    foreach ($results as $result) {
        $data[] = [
            'jahr' => $result->jahr,
            'monat' => $result->monat,
            'anzahl_anfragen' => $result->anzahl_anfragen,
        ];
    }

    return rest_ensure_response($data);
}
