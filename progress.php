<?php
session_start();  // Начинаем сессию

// Возвращаем прогресс, если он установлен
if (isset($_SESSION['progress'])) {
    echo json_encode([
        'status' => 'success',
        'progress' => $_SESSION['progress']
    ]);
} else {
    echo json_encode([
        'status' => 'error',
        'message' => 'Прогресс не найден'
    ]);
}
