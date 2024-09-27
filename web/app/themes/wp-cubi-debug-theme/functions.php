<?php

use OpenSpout\Writer\Common\Creator\WriterEntityFactory;

require_once __DIR__ . '/src/schema.php';
require_once __DIR__ . '/src/registrations.php';

add_action('save_post_registrations', 'send_registration_email', 10, 3);

function send_registration_email($post_id, $post, $update) {
    if (wp_is_post_autosave($post_id) || wp_is_post_revision($post_id) || $update) {
        return;
    }

    // Getting the registrant's email
    $registrant_email = get_post_meta($post_id, 'email', true);

    // Generating the PDF ticket
    $pdf_ticket = generate_pdf_ticket($post_id);

    // Prepare and send the email
    $subject = 'Your Event Ticket';
    $message = 'Thank you for registering. Your ticket is attached.';
    $headers = ['Content-Type: text/html; charset=UTF-8'];
    $attachments = [$pdf_ticket];

    // Send email
    wp_mail($registrant_email, $subject, $message, $headers, $attachments);
}

add_action('admin_post_export_registrations', 'handle_export_registrations');

function handle_export_registrations() {
    if (!current_user_can('manage_options')) {
        wp_die('Unauthorized user');
    }
    
    $event_id = intval($_GET['event_id']);
    
    // Fetch registrants for the event
    $registrants = get_posts([
        'post_type' => 'registrations',
        'meta_query' => [
            ['key' => 'registration_event_id', 'value' => $event_id],
        ],
        'numberposts' => -1,
    ]);
    
    // Generate the Excel file using OpenSpout
    generate_excel_file($registrants);
    exit;
}

function generate_excel_file($registrants) {
    $writer = WriterEntityFactory::createXLSXWriter();
    $writer->openToBrowser('registrants.xlsx');

    // Write header
    $header_row = WriterEntityFactory::createRowFromArray(['Name', 'Email']);
    $writer->addRow($header_row);

    // Write data rows
    foreach ($registrants as $registrant) {
        $name = get_post_meta($registrant->ID, 'name', true);
        $email = get_post_meta($registrant->ID, 'email', true);
        $row = WriterEntityFactory::createRowFromArray([$name, $email]);
        $writer->addRow($row);
    }

    $writer->close();
}