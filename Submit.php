<?php
// Load PHPMailer and PhpSpreadsheet libraries using Composer autoload
require 'vendor/autoload.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Collect form data
    $name = $_POST['name'] ?? '';
    $email = $_POST['email'] ?? '';
    $projectType = $_POST['project-type'] ?? '';
    $budget = $_POST['budget'] ?? '';
    $details = $_POST['details'] ?? '';
    $uploadedFile = $_FILES['reference-image'] ?? null;

    // Owner's email (change this to the recipient's email)
    $ownerEmail = 'owner@example.com';

    // Create Excel file with order details
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Order Details');

    // Add data to the Excel file
    $sheet->setCellValue('A1', 'Name')->setCellValue('B1', $name);
    $sheet->setCellValue('A2', 'Email')->setCellValue('B2', $email);
    $sheet->setCellValue('A3', 'Project Type')->setCellValue('B3', $projectType);
    $sheet->setCellValue('A4', 'Budget')->setCellValue('B4', $budget);
    $sheet->setCellValue('A5', 'Details')->setCellValue('B5', $details);

    // Save the Excel file to a temporary location
    $tempFile = tempnam(sys_get_temp_dir(), 'Order') . '.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save($tempFile);

    // Setup PHPMailer to send email
    $mail = new PHPMailer(true);

    try {
        // Server settings
        $mail->isSMTP();
        $mail->Host = 'smtp.example.com'; // SMTP server
        $mail->SMTPAuth = true;
        $mail->Username = 'your_email@example.com'; // SMTP username
        $mail->Password = 'your_password'; // SMTP password
        $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
        $mail->Port = 587;

        // Recipient and sender
        $mail->setFrom('your_email@example.com', 'Website Orders');
        $mail->addAddress($ownerEmail); // Owner's email

        // Add the Excel file attachment
        $mail->addAttachment($tempFile, 'Order_Details.xlsx');

        // Add uploaded file as attachment if it exists
        if ($uploadedFile && $uploadedFile['error'] === UPLOAD_ERR_OK) {
            $mail->addAttachment($uploadedFile['tmp_name'], $uploadedFile['name']);
        }

        // Email content
        $mail->isHTML(true);
        $mail->Subject = 'New Website Order';
        $mail->Body = "
            <h1>New Website Order</h1>
            <p><strong>Name:</strong> {$name}</p>
            <p><strong>Email:</strong> {$email}</p>
            <p><strong>Project Type:</strong> {$projectType}</p>
            <p><strong>Budget:</strong> {$budget}</p>
            <p><strong>Details:</strong> {$details}</p>
        ";

        // Send the email
        $mail->send();

        // Success response
        echo "Order submitted successfully. We will get back to you soon.";
    } catch (Exception $e) {
        echo "There was an error sending the email: {$mail->ErrorInfo}";
    } finally {
        // Clean up the temporary file
        unlink($tempFile);
    }
} else {
    echo "Invalid request.";
}
