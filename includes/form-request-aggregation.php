<?php
// Aggregationsskript (form-request-aggregation.php)

// Funktion zum Erstellen der neuen Tabelle für aggregierte Daten
function efs_create_aggregated_requests_table() {
    global $wpdb;
    $table_name = $wpdb->prefix . 'form_request_counts';

    if ($wpdb->get_var("SHOW TABLES LIKE '$table_name'") != $table_name) {
        $charset_collate = $wpdb->get_charset_collate();
        $sql = "CREATE TABLE $table_name (
            id bigint(20) UNSIGNED NOT NULL AUTO_INCREMENT,
            form_name varchar(255) NOT NULL,
            year int(4) NOT NULL,
            month int(2) NOT NULL,
            request_count int(11) NOT NULL,
            PRIMARY KEY (id),
            UNIQUE KEY form_year_month (form_name, year, month)
        ) $charset_collate;";

        require_once(ABSPATH . 'wp-admin/includes/upgrade.php');
        dbDelta($sql);
    }
}
add_action('after_setup_theme', 'efs_create_aggregated_requests_table');

// Funktion zum Aggregieren und Speichern der Anfragenanzahl
function efs_aggregate_form_requests() {
    global $wpdb;
    $table_name = $wpdb->prefix . 'e_submissions'; // Originale Tabelle mit Anfragen
    $count_table = $wpdb->prefix . 'form_request_counts'; // Neue Tabelle für aggregierte Daten

    $sql = "SELECT form_name, YEAR(created_at_gmt) AS jahr, MONTH(created_at_gmt) AS monat, COUNT(*) AS anzahl 
            FROM $table_name
            WHERE `status` NOT LIKE '%trash%'
            GROUP BY form_name, jahr, monat";

    $results = $wpdb->get_results($sql);

    if (!empty($results)) {
        foreach ($results as $result) {
            // Datensatz in der neuen Tabelle aktualisieren oder einfügen
            $wpdb->replace(
                $count_table,
                array(
                    'form_name' => $result->form_name,
                    'year' => $result->jahr,
                    'month' => $result->monat,
                    'request_count' => $result->anzahl,
                ),
                array(
                    '%s', '%d', '%d', '%d'
                )
            );
        }
    }
}

// Cron-Job für tägliche Aggregation einrichten
if (!wp_next_scheduled('efs_aggregate_form_requests_daily')) {
    wp_schedule_event(time(), 'daily', 'efs_aggregate_form_requests_daily');
}
add_action('efs_aggregate_form_requests_daily', 'efs_aggregate_form_requests');
