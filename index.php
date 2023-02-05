<?php
/**
 * Plugin Name: WooCommerce Order to Excel
 * Description: Replaces WooCommerce order emails with an Excel file attachment
 * Version: 1.0
 * Author: Marin
 * Author URI: monsite.com
 * License: GPL2
 */

use PHPExcel;
use PHPExcel_IOFactory;
// Ajouter un filtre pour remplacer l'email de commande par un fichier Excel

add_filter( 'woocommerce_email_attachments', 'replace_order_email_with_excel', 10, 3 );

function replace_order_email_with_excel( $attachments, $email_id, $order ) {
  if ( 'new_order' === $email_id && $order ) {
    $excel_file = generate_excel_from_order( $order );
    $attachments[] = $excel_file;
  }
  return $attachments;
}

function generer_excel_a_partir_de_commande( $commande ) {
    // Répertoire de sauvegarde
    $dir = './backup/';
  
    // Vérifier si le répertoire existe et est accessible en écriture
    if (!is_dir($dir)) {
      // Répertoire n'existe pas, le créer
      if (!mkdir($dir, 0777, true)) {
        trigger_error('Impossible de créer le répertoire de sauvegarde', E_USER_WARNING);
        return false;
      }
    }
    if (!is_writable($dir)) {
      // Répertoire n'est pas accessible en écriture
      trigger_error('Répertoire de sauvegarde non accessible en écriture', E_USER_WARNING);
      return false;
    }

function generate_excel_from_order( $order ) {
  $locale = get_locale();
  switch ( $locale ) {
    case 'fr_FR':
      $template_file = 'template-fr.xlsx';
      break;
    case 'zh_CN':
      $template_file = 'template-zh.xlsx';
      break;
    case 'en_US':
    default:
      $template_file = 'template-en.xlsx';
      break;
  }

  // Load the template file
  $excel = PHPExcel_IOFactory::load( $template_file );

  // Get the first worksheet
  $sheet = $excel->getSheet(0);

  // Set the order number
  $sheet->setCellValue( 'A1', $order->get_order_number() );

  // Set the order date
  $sheet->setCellValue( 'B1', $order->get_date_created()->format( 'Y-m-d' ) );

  // Set the customer information
  $sheet->setCellValue( 'A3', $order->get_billing_first_name() . ' ' . $order->get_billing_last_name() );
  $sheet->setCellValue( 'A4', $order->get_billing_email() );
  $sheet->setCellValue( 'A5', $order->get_billing_phone() );

  // Set the product information
  $row = 7;
  foreach ( $order->get_items() as $item ) {
    $sheet->setCellValue( 'A' . $row, $item->get_name() );
    $sheet->setCellValue( 'B' . $row, $item->get_quantity() );
    $sheet->setCellValue( 'C' . $row, $item->get_total() );
    // Add the serial number
    $sheet->setCellValue( 'D' . $row, get_serial_number( $item ) );
    $row++;
  }
  // Save the generated Excel file
  $file_path = $dir . 'order-' . $order->get_order_number() . '.xlsx';
  $writer = PHPExcel_IOFactory::createWriter( $excel, 'Excel2007' );
  
  try {
    $writer->save( $file_path );
  } catch (Exception $e) {
    // Handle error: unable to save the file
    trigger_error('Unable to save file: ' . $e->getMessage(), E_USER_WARNING);
    return false;
  }

  return $file_path;
}
