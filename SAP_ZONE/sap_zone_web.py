# -*- coding: utf-8 -*-
"""
SAP ZONE Modern Web UI Application (v1.6)
Giao diện Web App hiện đại (HTML5/CSS3/Tailwind-Glassmorphism Style)
Chạy trên Edge Chromium WebView2 hoặc Local Web App Mode.
Kết nối SAP chuẩn SAP_EMAIL_TEAMS_BRIDGE v1.1.12.
"""

import os
import sys
import json
import time
import threading
import subprocess
from datetime import datetime
from typing import Dict, Any, List

# Core logic
from sap_zone_core import (
    ConfigManager,
    ZoneManager,
    SapConnector,
    SapZonePipeline,
    ExcelResultExporter,
    MonStockService,
    parse_paste_data,
    get_base_dir
)

try:
    import webview
except ImportError:
    webview = None


# ==============================================================================
# HTML / CSS / JAVASCRIPT FRONTEND (Giao diện Web App Hiện Đại)
# ==============================================================================

HTML_UI = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SAP ZONE — Tự Động Ra Đơn & In Phiếu (v1.6)</title>
    <style>
        :root {
            --primary: #2563eb;
            --primary-dark: #1d4ed8;
            --primary-light: #eff6ff;
            --success: #16a34a;
            --success-dark: #15803d;
            --success-light: #ecfdf5;
            --danger: #dc2626;
            --warning: #d97706;
            --slate-50: #f8fafc;
            --slate-100: #f1f5f9;
            --slate-200: #e2e8f0;
            --slate-300: #cbd5e1;
            --slate-400: #94a3b8;
            --slate-600: #475569;
            --slate-700: #334155;
            --slate-800: #1e293b;
            --slate-900: #0f172a;
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 16px;
            --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
            --shadow-md: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
            --shadow-lg: 0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            -webkit-user-select: none;
            user-select: none;
        }

        input, textarea {
            -webkit-user-select: text;
            user-select: text;
        }

        body {
            background-color: var(--slate-100);
            color: var(--slate-900);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            overflow-x: hidden;
        }

        /* 1. HEADER BAR */
        .app-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 55%, #2563eb 100%);
            color: white;
            padding: 14px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
            position: sticky;
            top: 0;
            z-index: 50;
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .header-logo {
            width: 38px;
            height: 38px;
            background: rgba(255, 255, 255, 0.15);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.25);
            border-radius: var(--radius-md);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }

        .header-text h1 {
            font-size: 17px;
            font-weight: 700;
            letter-spacing: -0.3px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .badge-version {
            font-size: 11px;
            font-weight: 600;
            background: rgba(255, 255, 255, 0.2);
            padding: 2px 8px;
            border-radius: 999px;
            border: 1px solid rgba(255, 255, 255, 0.3);
        }

        .header-text p {
            font-size: 12px;
            color: #93c5fd;
            margin-top: 2px;
        }

        .header-right {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .sap-status-pill {
            display: flex;
            align-items: center;
            gap: 8px;
            background: rgba(15, 23, 42, 0.5);
            backdrop-filter: blur(6px);
            border: 1px solid rgba(255, 255, 255, 0.2);
            padding: 6px 14px;
            border-radius: 999px;
            font-size: 12px;
            font-weight: 600;
            color: #ffffff;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .sap-status-pill:hover {
            background: rgba(15, 23, 42, 0.7);
            border-color: rgba(255, 255, 255, 0.4);
        }

        .pulse-dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background-color: #22c55e;
            box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
            animation: pulse 1.8s infinite;
        }

        .pulse-dot.offline {
            background-color: #ef4444;
            box-shadow: none;
            animation: none;
        }

        @keyframes pulse {
            0% {
                transform: scale(0.95);
                box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7);
            }
            70% {
                transform: scale(1);
                box-shadow: 0 0 0 6px rgba(34, 197, 94, 0);
            }
            100% {
                transform: scale(0.95);
                box-shadow: 0 0 0 0 rgba(34, 197, 94, 0);
            }
        }

        /* 2. TABS NAVIGATION */
        .tab-bar-container {
            background: #ffffff;
            border-bottom: 1px solid var(--slate-200);
            padding: 0 24px;
            display: flex;
            align-items: center;
            justify-content: flex-start;
        }

        .tab-btn {
            padding: 12px 20px;
            font-size: 13.5px;
            font-weight: 600;
            color: var(--slate-600);
            background: transparent;
            border: none;
            border-bottom: 2px solid transparent;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .tab-btn:hover {
            color: var(--primary);
        }

        .tab-btn.active {
            color: var(--primary);
            border-bottom-color: var(--primary);
        }

        /* 3. MAIN CONTENT CONTAINER */
        .main-container {
            flex: 1;
            padding: 20px 24px;
            max-width: 1280px;
            width: 100%;
            margin: 0 auto;
        }

        .tab-page {
            display: none;
        }

        .tab-page.active {
            display: block;
        }

        /* 4. CARDS & SECTIONS */
        .card {
            background: #ffffff;
            border: 1px solid var(--slate-200);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            margin-bottom: 20px;
            overflow: hidden;
            transition: border-color 0.2s ease;
        }

        .card:hover {
            border-color: var(--slate-300);
        }

        .card-header {
            padding: 14px 20px;
            border-bottom: 1px solid var(--slate-200);
            background: var(--slate-50);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .card-header h2 {
            font-size: 14.5px;
            font-weight: 700;
            color: var(--slate-800);
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .card-header-actions {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .card-body {
            padding: 18px 20px;
        }

        /* 5. INPUT AREA */
        .textarea-wrapper {
            position: relative;
        }

        textarea.input-box {
            width: 100%;
            height: 120px;
            padding: 12px 14px;
            font-family: "Cascadia Code", "Fira Code", Consolas, monospace;
            font-size: 13px;
            line-height: 1.5;
            color: var(--slate-900);
            background: var(--slate-50);
            border: 1.5px solid var(--slate-200);
            border-radius: var(--radius-md);
            resize: vertical;
            outline: none;
            transition: all 0.2s ease;
        }

        textarea.input-box:focus {
            background: #ffffff;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
        }

        .input-meta-bar {
            margin-top: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 12px;
            color: var(--slate-600);
        }

        .meta-tag {
            background: var(--slate-100);
            border: 1px solid var(--slate-200);
            padding: 3px 10px;
            border-radius: 999px;
            font-weight: 600;
            color: var(--slate-700);
        }

        /* 6. BUTTONS */
        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            padding: 8px 14px;
            font-size: 13px;
            font-weight: 600;
            border-radius: var(--radius-sm);
            border: 1px solid transparent;
            cursor: pointer;
            transition: all 0.15s ease;
        }

        .btn-sm {
            padding: 5px 10px;
            font-size: 12px;
        }

        .btn-primary {
            background: var(--primary);
            color: #ffffff;
        }

        .btn-primary:hover {
            background: var(--primary-dark);
        }

        .btn-secondary {
            background: var(--slate-100);
            color: var(--slate-700);
            border-color: var(--slate-200);
        }

        .btn-secondary:hover {
            background: var(--slate-200);
            color: var(--slate-900);
        }

        .btn-success {
            background: var(--success);
            color: #ffffff;
        }

        .btn-success:hover {
            background: var(--success-dark);
        }

        .btn-danger {
            background: var(--danger);
            color: #ffffff;
        }

        .btn-danger:hover {
            background: #b91c1c;
        }

        /* Nút Chạy Hero to nổi bật */
        .hero-action-bar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 14px 20px;
            background: #ffffff;
            border: 1px solid var(--slate-200);
            border-radius: var(--radius-lg);
            margin-bottom: 20px;
            box-shadow: var(--shadow-sm);
        }

        .mock-toggle {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            font-weight: 500;
            color: var(--slate-700);
            cursor: pointer;
        }

        .btn-hero-run {
            padding: 12px 28px;
            font-size: 15px;
            font-weight: 700;
            letter-spacing: -0.2px;
            border-radius: var(--radius-md);
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            color: white;
            border: none;
            box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35);
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 10px;
            transition: all 0.2s ease;
        }

        .btn-hero-run:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(37, 99, 235, 0.45);
        }

        .btn-hero-run:active {
            transform: translateY(1px);
        }

        .btn-hero-run.running {
            background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%);
            box-shadow: 0 4px 14px rgba(220, 38, 38, 0.35);
        }

        /* 7. PROGRESS BAR & STEP TICKER */
        .progress-card {
            background: #ffffff;
            border: 1px solid var(--slate-200);
            border-radius: var(--radius-lg);
            padding: 16px 20px;
            margin-bottom: 20px;
            display: none;
        }

        .progress-card.active {
            display: block;
        }

        .progress-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
            font-size: 13px;
            font-weight: 600;
        }

        .progress-track {
            width: 100%;
            height: 10px;
            background: var(--slate-100);
            border-radius: 999px;
            overflow: hidden;
            border: 1px solid var(--slate-200);
            margin-bottom: 12px;
        }

        .progress-fill {
            height: 100%;
            width: 0%;
            background: linear-gradient(90deg, #2563eb, #3b82f6);
            border-radius: 999px;
            transition: width 0.3s ease;
        }

        .step-stepper {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-top: 10px;
            padding-top: 10px;
            border-top: 1px dashed var(--slate-200);
        }

        .step-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 11.5px;
            font-weight: 500;
            color: var(--slate-400);
        }

        .step-item.active {
            color: var(--primary);
            font-weight: 700;
        }

        .step-item.done {
            color: var(--success);
        }

        .step-icon {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: var(--slate-200);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
        }

        .step-item.active .step-icon {
            background: var(--primary);
            color: white;
        }

        .step-item.done .step-icon {
            background: var(--success);
            color: white;
        }

        /* 8. RESULT TABLE */
        .table-responsive {
            width: 100%;
            overflow-x: auto;
            max-height: 380px;
            overflow-y: auto;
            border-radius: var(--radius-md);
            border: 1px solid var(--slate-200);
        }

        table.data-table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 13px;
        }

        table.data-table thead th {
            position: sticky;
            top: 0;
            background: #f8fafc;
            color: var(--slate-700);
            font-weight: 700;
            padding: 10px 14px;
            border-bottom: 1.5px solid var(--slate-200);
            white-space: nowrap;
            z-index: 10;
        }

        table.data-table tbody td {
            padding: 10px 14px;
            border-bottom: 1px solid var(--slate-200);
            color: var(--slate-900);
        }

        table.data-table tbody tr:hover {
            background-color: #f1f5f9;
        }

        .ob-pill {
            font-family: monospace;
            font-weight: 700;
            color: #1e40af;
            background: #dbeafe;
            padding: 3px 8px;
            border-radius: 6px;
            display: inline-block;
        }

        .mat-code {
            font-family: monospace;
            font-weight: 700;
            color: var(--slate-900);
        }

        .who-pill {
            font-family: monospace;
            font-weight: 600;
            color: #065f46;
            background: #d1fae5;
            padding: 3px 8px;
            border-radius: 6px;
            display: inline-block;
        }

        .empty-state {
            text-align: center;
            padding: 40px 20px;
            color: var(--slate-400);
        }

        .empty-state-icon {
            font-size: 36px;
            margin-bottom: 8px;
            opacity: 0.6;
        }

        /* 9. SETTINGS TAB (GRID) */
        .settings-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
            gap: 20px;
        }

        .form-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
            font-size: 13px;
        }

        .form-row label {
            font-weight: 500;
            color: var(--slate-700);
            flex: 1;
        }

        .form-row input[type="text"],
        .form-row input[type="number"],
        .form-row select {
            width: 160px;
            padding: 6px 10px;
            border: 1px solid var(--slate-300);
            border-radius: var(--radius-sm);
            font-size: 13px;
            outline: none;
            background: #ffffff;
        }

        .form-row input:focus, .form-row select:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.15);
        }

        /* 10. TOAST NOTIFICATION */
        .toast-container {
            position: fixed;
            bottom: 24px;
            right: 24px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 8px;
            pointer-events: none;
        }

        .toast {
            background: #0f172a;
            color: white;
            padding: 12px 18px;
            border-radius: var(--radius-md);
            font-size: 13px;
            font-weight: 500;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.25);
            display: flex;
            align-items: center;
            gap: 10px;
            opacity: 0;
            transform: translateY(20px);
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            pointer-events: auto;
        }

        .toast.show {
            opacity: 1;
            transform: translateY(0);
        }

        .toast.success {
            border-left: 4px solid #22c55e;
        }

        .toast.error {
            border-left: 4px solid #ef4444;
        }

        .toast.info {
            border-left: 4px solid #3b82f6;
        }
    </style>
</head>
<body>

    <!-- 1. HEADER BAR -->
    <header class="app-header">
        <div class="header-left">
            <div class="header-logo">⚡</div>
            <div class="header-text">
                <h1>SAP ZONE <span class="badge-version">v1.6 Pro</span></h1>
                <p>EWM APFL • FIFO/FEFO Allocation • MIGO 311 • WHO F00112</p>
            </div>
        </div>
        <div class="header-right">
            <div class="sap-status-pill" id="btn-check-sap" onclick="checkSapStatus()" title="Bấm để kiểm tra lại phiên SAP">
                <div class="pulse-dot" id="sap-dot"></div>
                <span id="sap-status-text">Đang kiểm tra SAP...</span>
                <span>🔄</span>
            </div>
        </div>
    </header>

    <!-- 2. TABS NAVIGATION -->
    <div class="tab-bar-container">
        <button class="tab-btn active" id="tab-nav-main" onclick="switchTab('main')">
            <span>🚀</span> VẬN HÀNH CHÍNH (1-Click Run)
        </button>
        <button class="tab-btn" id="tab-nav-config" onclick="switchTab('config')">
            <span>⚙️</span> CẤU HÌNH HỆ THỐNG
        </button>
    </div>

    <!-- 3. MAIN CONTENT -->
    <main class="main-container">

        <!-- ============================================================== -->
        <!-- TAB 1: VẬN HÀNH CHÍNH                                         -->
        <!-- ============================================================== -->
        <section id="page-main" class="tab-page active">

            <!-- Card 1: Nhập liệu -->
            <div class="card">
                <div class="card-header">
                    <h2><span>📋</span> 1. Dán Dữ Liệu Yêu Cầu Ra Đơn (Từ Excel)</h2>
                    <div class="card-header-actions">
                        <button class="btn btn-secondary btn-sm" onclick="pasteClipboard()">📋 Dán từ Clipboard</button>
                        <button class="btn btn-secondary btn-sm" onclick="loadSampleData()">📝 Nạp Mẫu Thử</button>
                        <button class="btn btn-secondary btn-sm" onclick="clearInput()">🗑 Xóa</button>
                    </div>
                </div>
                <div class="card-body">
                    <div class="textarea-wrapper">
                        <textarea class="input-box" id="raw-input" placeholder="Dán danh sách mã hàng từ Excel vào đây...&#10;Định dạng cột: [Mã hàng]  [Số lượng]  [Trạm / Người nhận]&#10;Ví dụ:&#10;BIW32211725AA	300	Hùng/Long&#10;BIW32211727AA	300	Hùng/Long" oninput="updateInputStats()"></textarea>
                    </div>
                    <div class="input-meta-bar">
                        <div>
                            <span>Trạng thái: </span>
                            <span class="meta-tag" id="input-stats-badge">0 dòng hợp lệ</span>
                        </div>
                        <span style="font-style: italic; color: var(--slate-400);">Hệ thống tự nhận diện khoảng trắng, dấu Tab và tự động bỏ dòng tiêu đề.</span>
                    </div>
                </div>
            </div>

            <!-- Hero Action Bar -->
            <div class="hero-action-bar">
                <label class="mock-toggle">
                    <input type="checkbox" id="chk-mock-mode">
                    <span>Chạy mô phỏng (Test Mock Data không cần SAP)</span>
                </label>
                <button class="btn-hero-run" id="btn-run" onclick="toggleRun()">
                    <span id="run-btn-icon">▶</span>
                    <span id="run-btn-text">BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU</span>
                </button>
            </div>

            <!-- Progress & Status Card -->
            <div class="progress-card" id="progress-card">
                <div class="progress-header">
                    <span id="progress-status-msg" style="color: var(--primary);">Đang khởi tạo tiến trình...</span>
                    <span id="progress-pct-badge" class="meta-tag" style="color: var(--primary); background: var(--primary-light);">0%</span>
                </div>
                <div class="progress-track">
                    <div class="progress-fill" id="progress-fill"></div>
                </div>
                <div class="step-stepper">
                    <div class="step-item" id="step-1"><div class="step-icon">1</div> SAP GUI</div>
                    <div class="step-item" id="step-2"><div class="step-icon">2</div> Đọc Tồn MON</div>
                    <div class="step-item" id="step-3"><div class="step-icon">3</div> Gom Nhóm Zone</div>
                    <div class="step-item" id="step-4"><div class="step-icon">4</div> MIGO 311</div>
                    <div class="step-item" id="step-5"><div class="step-icon">5</div> Tra Cứu WHO</div>
                    <div class="step-item" id="step-6"><div class="step-icon">6</div> In Phiếu PDF</div>
                </div>
            </div>

            <!-- Card 2: Bảng kết quả 3 cột -->
            <div class="card">
                <div class="card-header">
                    <h2><span>📊</span> 2. Bảng Kết Quả Theo Dõi (3 Cột Chuẩn)</h2>
                    <div class="card-header-actions">
                        <span class="meta-tag" id="results-count-badge">Tổng cộng: 0 dòng</span>
                        <button class="btn btn-success btn-sm" onclick="copy3Columns()" id="btn-copy-3col">
                            📋 Sao Chép 3 Cột (Excel)
                        </button>
                        <button class="btn btn-secondary btn-sm" onclick="exportExcel()">📊 Xuất Excel</button>
                        <button class="btn btn-secondary btn-sm" onclick="openPdfFolder()">📁 Mở Thư Mục PDF</button>
                    </div>
                </div>
                <div class="card-body" style="padding: 12px;">
                    <div class="table-responsive">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th style="width: 25%; text-align: center;">Số OB</th>
                                    <th style="width: 35%;">Mã Hàng (Material)</th>
                                    <th style="width: 40%;">Số Đơn (WHO)</th>
                                </tr>
                            </thead>
                            <tbody id="table-results-body">
                                <tr>
                                    <td colspan="3">
                                        <div class="empty-state">
                                            <div class="empty-state-icon">📋</div>
                                            <p>Chưa có kết quả ra đơn. Hãy dán dữ liệu và bấm <b>Bắt Đầu Chạy</b>.</p>
                                        </div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

        </section>

        <!-- ============================================================== -->
        <!-- TAB 2: CẤU HÌNH HỆ THỐNG                                      -->
        <!-- ============================================================== -->
        <section id="page-config" class="tab-page">
            <div class="settings-grid">

                <!-- Khung 1: SAP GUI -->
                <div class="card">
                    <div class="card-header">
                        <h2><span>🖥️</span> 1. Kết Nối SAP GUI</h2>
                        <button class="btn btn-primary btn-sm" onclick="checkSapStatus()">🔌 Thử Kết Nối</button>
                    </div>
                    <div class="card-body">
                        <div class="form-row">
                            <label>Hệ thống (System):</label>
                            <input type="text" id="cfg-sap-sys" value="PSV">
                        </div>
                        <div class="form-row">
                            <label>Client:</label>
                            <input type="text" id="cfg-sap-client" value="908">
                        </div>
                        <div class="form-row">
                            <label>Session Index (Vị trí tab):</label>
                            <input type="number" id="cfg-sap-idx" value="0" min="0" max="5">
                        </div>
                        <p style="font-size: 11.5px; color: var(--slate-400); margin-top: 6px;">
                            * Hệ thống tự động nhận diện danh tính phiên thực tế (PSV/PRD, 908/300) khi kết nối.
                        </p>
                    </div>
                </div>

                <!-- Khung 2: Kho EWM & MIGO -->
                <div class="card">
                    <div class="card-header">
                        <h2><span>🏭</span> 2. Tham Số Kho EWM & Ra Đơn</h2>
                    </div>
                    <div class="card-body">
                        <div class="form-row">
                            <label>Mã kho (Warehouse):</label>
                            <input type="text" id="cfg-whse" value="VF02">
                        </div>
                        <div class="form-row">
                            <label>Storage Type đọc tồn:</label>
                            <input type="text" id="cfg-storage" value="APFL">
                        </div>
                        <div class="form-row">
                            <label>Chủ hàng (Owner):</label>
                            <input type="text" id="cfg-owner" value="3093">
                        </div>
                        <div class="form-row">
                            <label>Plant MIGO:</label>
                            <input type="text" id="cfg-plant" value="3093">
                        </div>
                        <div class="form-row">
                            <label>SLoc MIGO:</label>
                            <input type="text" id="cfg-sloc" value="3093">
                        </div>
                    </div>
                </div>

                <!-- Khung 3: Gom nhóm & Phân bổ -->
                <div class="card">
                    <div class="card-header">
                        <h2><span>📦</span> 3. Quy Tắc Gom Nhóm & FIFO</h2>
                    </div>
                    <div class="card-body">
                        <div class="form-row">
                            <label>Số mã tối đa / OB:</label>
                            <input type="number" id="cfg-max-mats" value="5" min="1" max="20">
                        </div>
                        <div class="form-row">
                            <label>Số dòng tối đa / OB:</label>
                            <input type="number" id="cfg-max-lines" value="20" min="1" max="50">
                        </div>
                        <div class="form-row">
                            <label>Chiến lược cấp phát:</label>
                            <select id="cfg-strategy">
                                <option value="FIFO">FIFO (Nhập trước xuất trước)</option>
                                <option value="FEFO">FEFO (Hạn dùng trước)</option>
                            </select>
                        </div>
                    </div>
                </div>

                <!-- Khung 4: In ấn & PDF -->
                <div class="card">
                    <div class="card-header">
                        <h2><span>🖨️</span> 4. Cấu Hình In Phiếu SAP & PDF</h2>
                    </div>
                    <div class="card-body">
                        <div class="form-row">
                            <label>Số bản in SAP (Copies):</label>
                            <input type="number" id="cfg-copies" value="2" min="1" max="5">
                        </div>
                        <div class="form-row">
                            <label>Tự động lưu PDF:</label>
                            <input type="checkbox" id="cfg-auto-print" checked style="width: auto;">
                        </div>
                        <div class="form-row">
                            <label>Thư mục lưu PDF:</label>
                            <input type="text" id="cfg-pdf-dir" style="width: 100%; margin-top: 6px;">
                        </div>
                    </div>
                </div>

                <!-- Khung 5: Zone Mapping -->
                <div class="card" style="grid-column: 1 / -1;">
                    <div class="card-header">
                        <h2><span>🗺️</span> 5. Dữ Liệu Ánh Xạ Trạm & Zone (zone_map.json)</h2>
                        <div class="card-header-actions">
                            <button class="btn btn-secondary btn-sm" onclick="openZoneFile()">📝 Xem File JSON</button>
                            <button class="btn btn-secondary btn-sm" onclick="reloadZoneMap()">🔄 Nạp Lại Dữ Liệu</button>
                        </div>
                    </div>
                    <div class="card-body">
                        <p id="zone-status-text" style="font-size: 13px; color: var(--slate-700);">
                            Đang tải thông tin quy tắc ánh xạ...
                        </p>
                    </div>
                </div>

            </div>

            <!-- Nút Lưu Cấu Hình -->
            <div style="margin-top: 16px; display: flex; justify-content: flex-end;">
                <button class="btn btn-success" style="padding: 12px 24px; font-size: 14px;" onclick="saveConfig()">
                    💾 LƯU CẤU HÌNH HỆ THỐNG
                </button>
            </div>
        </section>

    </main>

    <!-- Toast Notifications Container -->
    <div class="toast-container" id="toast-container"></div>

    <!-- JAVASCRIPT LOGIC -->
    <script>
        let currentResults = [];
        let isRunning = false;
        let pollTimer = null;

        // --- 1. TOAST NOTIFICATION ---
        function showToast(message, type = 'info') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            const icon = type === 'success' ? '✅' : (type === 'error' ? '❌' : 'ℹ️');
            toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
            container.appendChild(toast);
            setTimeout(() => toast.classList.add('show'), 20);
            setTimeout(() => {
                toast.classList.remove('show');
                setTimeout(() => toast.remove(), 300);
            }, 3500);
        }

        // --- 2. TAB SWITCHING ---
        function switchTab(tabId) {
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-page').forEach(page => page.classList.remove('active'));

            if (tabId === 'main') {
                document.getElementById('tab-nav-main').classList.add('active');
                document.getElementById('page-main').classList.add('active');
            } else {
                document.getElementById('tab-nav-config').classList.add('active');
                document.getElementById('page-config').classList.add('active');
            }
        }

        // --- 3. INPUT HELPERS ---
        function clearInput() {
            document.getElementById('raw-input').value = '';
            updateInputStats();
        }

        function loadSampleData() {
            const sample = "BIW32211725AA\t300\tHùng/Long\\nBIW32211727AA\t300\tHùng/Long";
            document.getElementById('raw-input').value = sample;
            updateInputStats();
            showToast('Đã nạp dữ liệu mẫu thử nghiệm!', 'info');
        }

        async function pasteClipboard() {
            try {
                if (window.pywebview && window.pywebview.api) {
                    const text = await window.pywebview.api.get_clipboard_text();
                    if (text) {
                        document.getElementById('raw-input').value = text;
                        updateInputStats();
                        showToast('Đã dán dữ liệu từ Clipboard!', 'success');
                        return;
                    }
                }
                const text = await navigator.clipboard.readText();
                if (text) {
                    document.getElementById('raw-input').value = text;
                    updateInputStats();
                    showToast('Đã dán dữ liệu từ Clipboard!', 'success');
                }
            } catch (err) {
                showToast('Không thể đọc Clipboard: ' + err, 'error');
            }
        }

        function updateInputStats() {
            const text = document.getElementById('raw-input').value.trim();
            if (!text) {
                document.getElementById('input-stats-badge').innerText = '0 dòng hợp lệ';
                return;
            }
            const lines = text.split('\\n').filter(l => l.trim().length > 0);
            let validCount = 0;
            let totalQty = 0;

            lines.forEach(l => {
                const parts = l.split(/[\\t,;\\s]+/).filter(p => p.length > 0);
                if (parts.length >= 2) {
                    const q = parseFloat(parts[1].replace(',', '.'));
                    if (!isNaN(q) && q > 0) {
                        validCount++;
                        totalQty += q;
                    }
                }
            });

            document.getElementById('input-stats-badge').innerText = 
                `${validCount} dòng hợp lệ (Tổng SL: ${totalQty.toLocaleString('vi-VN')})`;
        }

        // --- 4. SAP STATUS CHECK ---
        async function checkSapStatus() {
            const pillText = document.getElementById('sap-status-text');
            const dot = document.getElementById('sap-dot');
            pillText.innerText = 'Đang kiểm tra...';

            try {
                if (window.pywebview && window.pywebview.api) {
                    const res = await window.pywebview.api.get_sap_status();
                    if (res.ok) {
                        pillText.innerText = `SAP: ${res.system}/${res.client} (${res.user})`;
                        dot.className = 'pulse-dot';
                        showToast(`Kết nối SAP thành công: ${res.system}/${res.client} - User: ${res.user}`, 'success');
                    } else {
                        pillText.innerText = 'Chưa Kết Nối SAP';
                        dot.className = 'pulse-dot offline';
                        showToast('Chưa kết nối SAP: ' + (res.error || 'Vui lòng mở SAP Logon'), 'error');
                    }
                }
            } catch (ex) {
                pillText.innerText = 'Lỗi kết nối';
                dot.className = 'pulse-dot offline';
            }
        }

        // --- 5. RUN PIPELINE ---
        async function toggleRun() {
            if (isRunning) {
                // Request stop
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.stop_pipeline();
                }
                showToast('Đang yêu cầu dừng tiến trình...', 'warning');
                return;
            }

            const rawText = document.getElementById('raw-input').value.trim();
            if (!rawText) {
                showToast('Vui lòng dán danh sách mã hàng trước khi bấm Chạy!', 'error');
                return;
            }

            const isMock = document.getElementById('chk-mock-mode').checked;

            // UI update to running state
            isRunning = true;
            const btn = document.getElementById('btn-run');
            btn.classList.add('running');
            document.getElementById('run-btn-icon').innerText = '⏹';
            document.getElementById('run-btn-text').innerText = 'DỪNG LẠI';

            // Show progress
            document.getElementById('progress-card').classList.add('active');
            updateProgressUI(0, 'Bắt đầu xử lý...', 1);

            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.start_pipeline(rawText, isMock);
                    // Start polling status
                    startPolling();
                }
            } catch (ex) {
                stopRunningUI();
                showToast('Lỗi khởi chạy: ' + ex, 'error');
            }
        }

        function startPolling() {
            if (pollTimer) clearInterval(pollTimer);
            pollTimer = setInterval(async () => {
                if (!window.pywebview || !window.pywebview.api) return;
                try {
                    const prog = await window.pywebview.api.get_progress();
                    updateProgressUI(prog.percent, prog.status_text, prog.step_index);

                    if (!prog.is_running) {
                        clearInterval(pollTimer);
                        pollTimer = null;
                        stopRunningUI();

                        if (prog.success) {
                            currentResults = prog.records || [];
                            renderResults(currentResults);
                            showToast(`Hoàn tất thành công! Đã xử lý ${currentResults.length} dòng kết quả.`, 'success');
                        } else {
                            showToast('Thất bại: ' + (prog.error || 'Lỗi không xác định'), 'error');
                        }
                    }
                } catch (e) {
                    console.error('Poll err:', e);
                }
            }, 400);
        }

        function stopRunningUI() {
            isRunning = false;
            const btn = document.getElementById('btn-run');
            btn.classList.remove('running');
            document.getElementById('run-btn-icon').innerText = '▶';
            document.getElementById('run-btn-text').innerText = 'BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU';
        }

        function updateProgressUI(pct, msg, stepIdx = 1) {
            document.getElementById('progress-fill').style.width = pct + '%';
            document.getElementById('progress-pct-badge').innerText = pct + '%';
            document.getElementById('progress-status-msg').innerText = msg;

            // Stepper
            for (let i = 1; i <= 6; i++) {
                const el = document.getElementById('step-' + i);
                if (el) {
                    el.classList.remove('active', 'done');
                    if (i < stepIdx) el.classList.add('done');
                    else if (i === stepIdx) el.classList.add('active');
                }
            }
        }

        // --- 6. RENDER RESULTS ---
        function renderResults(records) {
            const tbody = document.getElementById('table-results-body');
            document.getElementById('results-count-badge').innerText = `Tổng cộng: ${records.length} dòng`;

            if (!records || records.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="3">
                            <div class="empty-state">
                                <div class="empty-state-icon">📋</div>
                                <p>Không có dòng kết quả nào.</p>
                            </div>
                        </td>
                    </tr>`;
                return;
            }

            let html = '';
            records.forEach(r => {
                html += `
                    <tr>
                        <td style="text-align: center;"><span class="ob-pill">${r.ob}</span></td>
                        <td><span class="mat-code">${r.material}</span></td>
                        <td><span class="who-pill">${r.who}</span></td>
                    </tr>`;
            });
            tbody.innerHTML = html;
        }

        // --- 7. ACTIONS (COPY, EXCEL, FOLDER) ---
        async function copy3Columns() {
            if (!currentResults || currentResults.length === 0) {
                showToast('Chưa có kết quả để sao chép!', 'warning');
                return;
            }

            const tsv = currentResults.map(r => `${r.ob}\\t${r.material}\\t${r.who}`).join('\\r\\n');

            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.copy_text(tsv);
                } else {
                    await navigator.clipboard.writeText(tsv);
                }
                showToast(`Đã sao chép ${currentResults.length} dòng vào Clipboard! Mở Excel và bấm Ctrl+V.`, 'success');
            } catch (ex) {
                showToast('Lỗi sao chép: ' + ex, 'error');
            }
        }

        async function exportExcel() {
            if (!currentResults || currentResults.length === 0) {
                showToast('Chưa có kết quả để xuất Excel!', 'warning');
                return;
            }
            try {
                if (window.pywebview && window.pywebview.api) {
                    const res = await window.pywebview.api.export_excel(currentResults);
                    if (res.ok) {
                        showToast('Đã xuất Excel: ' + res.path, 'success');
                    } else {
                        showToast('Lỗi xuất Excel: ' + res.error, 'error');
                    }
                }
            } catch (ex) {
                showToast('Lỗi xuất Excel: ' + ex, 'error');
            }
        }

        async function openPdfFolder() {
            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.open_pdf_folder();
                }
            } catch (ex) {
                showToast('Lỗi mở thư mục: ' + ex, 'error');
            }
        }

        // --- 8. CONFIG HELPERS ---
        async function loadConfigUI() {
            try {
                if (window.pywebview && window.pywebview.api) {
                    const cfg = await window.pywebview.api.get_config();
                    if (cfg) {
                        document.getElementById('cfg-sap-sys').value = cfg.sap_system || 'PSV';
                        document.getElementById('cfg-sap-client').value = cfg.sap_client || '908';
                        document.getElementById('cfg-sap-idx').value = cfg.sap_session_idx || 0;
                        document.getElementById('cfg-whse').value = cfg.ewm_warehouse || 'VF02';
                        document.getElementById('cfg-storage').value = cfg.ewm_storage_type || 'APFL';
                        document.getElementById('cfg-owner').value = cfg.ewm_owner || '3093';
                        document.getElementById('cfg-plant').value = cfg.plant || '3093';
                        document.getElementById('cfg-sloc').value = cfg.sloc || '3093';
                        document.getElementById('cfg-max-mats').value = cfg.max_materials_per_group || 5;
                        document.getElementById('cfg-max-lines').value = cfg.max_lines_per_group || 20;
                        document.getElementById('cfg-strategy').value = cfg.allocation_strategy || 'FIFO';
                        document.getElementById('cfg-copies').value = cfg.sap_print_copies || 2;
                        document.getElementById('cfg-auto-print').checked = !!cfg.auto_print_pdf;
                        document.getElementById('cfg-pdf-dir').value = cfg.pdf_output_folder || '';
                    }

                    const zoneInfo = await window.pywebview.api.get_zone_info();
                    document.getElementById('zone-status-text').innerText = 
                        `Hiện có: ${zoneInfo.count} quy tắc ánh xạ Trạm -> Zone -> Tổ trưởng từ file zone_map.json.`;
                }
            } catch (ex) {
                console.error('Config load err:', ex);
            }
        }

        async function saveConfig() {
            const cfg = {
                sap_system: document.getElementById('cfg-sap-sys').value.trim(),
                sap_client: document.getElementById('cfg-sap-client').value.trim(),
                sap_session_idx: parseInt(document.getElementById('cfg-sap-idx').value) || 0,
                ewm_warehouse: document.getElementById('cfg-whse').value.trim(),
                ewm_storage_type: document.getElementById('cfg-storage').value.trim(),
                ewm_owner: document.getElementById('cfg-owner').value.trim(),
                plant: document.getElementById('cfg-plant').value.trim(),
                sloc: document.getElementById('cfg-sloc').value.trim(),
                max_materials_per_group: parseInt(document.getElementById('cfg-max-mats').value) || 5,
                max_lines_per_group: parseInt(document.getElementById('cfg-max-lines').value) || 20,
                allocation_strategy: document.getElementById('cfg-strategy').value,
                sap_print_copies: parseInt(document.getElementById('cfg-copies').value) || 2,
                auto_print_pdf: document.getElementById('cfg-auto-print').checked,
                pdf_output_folder: document.getElementById('cfg-pdf-dir').value.trim()
            };

            try {
                if (window.pywebview && window.pywebview.api) {
                    const ok = await window.pywebview.api.save_config(cfg);
                    if (ok) {
                        showToast('Đã lưu cấu hình vào config.json!', 'success');
                    } else {
                        showToast('Không thể lưu file config.json!', 'error');
                    }
                }
            } catch (ex) {
                showToast('Lỗi lưu cấu hình: ' + ex, 'error');
            }
        }

        async function reloadZoneMap() {
            try {
                if (window.pywebview && window.pywebview.api) {
                    const res = await window.pywebview.api.reload_zone_map();
                    document.getElementById('zone-status-text').innerText = 
                        `Hiện có: ${res.count} quy tắc ánh xạ Trạm -> Zone -> Tổ trưởng.`;
                    showToast(`Đã nạp lại ${res.count} quy tắc Zone Mapping!`, 'success');
                }
            } catch (ex) {
                showToast('Lỗi nạp Zone: ' + ex, 'error');
            }
        }

        async function openZoneFile() {
            try {
                if (window.pywebview && window.pywebview.api) {
                    await window.pywebview.api.open_zone_file();
                }
            } catch (ex) {
                showToast('Lỗi mở file: ' + ex, 'error');
            }
        }

        // --- INIT ---
        window.addEventListener('pywebviewready', () => {
            checkSapStatus();
            loadConfigUI();
        });

        // Fallback init
        setTimeout(() => {
            if (window.pywebview) {
                checkSapStatus();
                loadConfigUI();
            }
        }, 1000);
    </script>
</body>
</html>
"""


# ==============================================================================
# PYTHON BACKEND API (Giao tiếp với PyWebView / Frontend)
# ==============================================================================

class SapZoneWebBackend:
    """API Backend cung cấp dịch vụ cho giao diện Web."""

    def __init__(self):
        self.config = ConfigManager.load_config()
        self.zone_mgr = ZoneManager()
        self.pipeline: Optional[SapZonePipeline] = None
        self.worker_thread: Optional[threading.Thread] = None

        # Trạng thái tiến trình
        self.progress_state = {
            "percent": 0,
            "status_text": "Sẵn sàng",
            "step_index": 1,
            "is_running": False,
            "success": False,
            "error": "",
            "records": []
        }

    def get_sap_status(self) -> Dict[str, Any]:
        """Kiểm tra và tự động bắt tay phiên SAP theo chuẩn SAP_EMAIL_TEAMS_BRIDGE v1.1.12."""
        s_idx = int(self.config.get("sap_session_idx", -1))
        res = SapConnector.test_connection(s_idx)
        return res

    def get_config(self) -> Dict[str, Any]:
        self.config = ConfigManager.load_config()
        return self.config

    def save_config(self, new_cfg: Dict[str, Any]) -> bool:
        ok = ConfigManager.save_config(new_cfg)
        if ok:
            self.config = new_cfg
        return ok

    def get_zone_info(self) -> Dict[str, Any]:
        return {"count": len(self.zone_mgr.mappings)}

    def reload_zone_map(self) -> Dict[str, Any]:
        self.zone_mgr.load_mappings()
        return {"count": len(self.zone_mgr.mappings)}

    def open_zone_file(self) -> bool:
        path = self.zone_mgr.get_path()
        if os.path.exists(path):
            try:
                os.startfile(path)
                return True
            except Exception:
                pass
        return False

    def get_clipboard_text(self) -> str:
        """Đọc văn bản từ Clipboard hệ điều hành."""
        try:
            import win32clipboard
            import win32con
            win32clipboard.OpenClipboard()
            text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return text or ""
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return ""

    def copy_text(self, text: str) -> bool:
        """Ghi văn bản vào Clipboard hệ điều hành."""
        return MonStockService.copy_to_clipboard(text)

    def open_pdf_folder(self) -> bool:
        folder = self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF"))
        os.makedirs(folder, exist_ok=True)
        try:
            os.startfile(folder)
            return True
        except Exception:
            return False

    def export_excel(self, records: List[Dict[str, str]]) -> Dict[str, Any]:
        default_dir = os.path.abspath("./OUTPUT_EXCEL")
        os.makedirs(default_dir, exist_ok=True)
        filename = f"KET_QUA_RA_DON_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(default_dir, filename)

        ok = ExcelResultExporter.export(records, file_path)
        if ok:
            try:
                os.startfile(file_path)
            except Exception:
                pass
            return {"ok": True, "path": file_path}
        else:
            return {"ok": False, "error": "Lỗi khi ghi file Excel"}

    def start_pipeline(self, raw_text: str, is_mock: bool = False) -> bool:
        if self.progress_state["is_running"]:
            return False

        requests = parse_paste_data(raw_text)
        if not requests:
            raise ValueError("Không nhận diện được dòng dữ liệu hợp lệ nào!")

        # Reset tiến trình
        self.progress_state = {
            "percent": 0,
            "status_text": "Khởi động...",
            "step_index": 1,
            "is_running": True,
            "success": False,
            "error": "",
            "records": []
        }

        self.pipeline = SapZonePipeline(self.config, self.zone_mgr)
        self.worker_thread = threading.Thread(
            target=self._worker_run,
            args=(requests, is_mock),
            daemon=True
        )
        self.worker_thread.start()
        return True

    def _worker_run(self, requests, is_mock):
        # Đảm bảo COM được khởi tạo trong Thread con
        SapConnector.init_thread_com()

        def prog_cb(pct):
            self.progress_state["percent"] = pct
            if pct <= 10:
                self.progress_state["step_index"] = 1
            elif pct <= 30:
                self.progress_state["step_index"] = 2
            elif pct <= 50:
                self.progress_state["step_index"] = 3
            elif pct <= 70:
                self.progress_state["step_index"] = 4
            elif pct <= 85:
                self.progress_state["step_index"] = 5
            else:
                self.progress_state["step_index"] = 6

        def status_cb(msg):
            self.progress_state["status_text"] = msg

        try:
            summary = self.pipeline.run(
                requests=requests,
                progress_cb=prog_cb,
                status_cb=status_cb,
                is_mock=is_mock
            )
            self.progress_state["is_running"] = False
            self.progress_state["success"] = summary.get("success", False)
            self.progress_state["records"] = summary.get("records", [])
            if not summary.get("success"):
                self.progress_state["error"] = summary.get("message", "Lỗi xử lý")
        except Exception as ex:
            self.progress_state["is_running"] = False
            self.progress_state["success"] = False
            self.progress_state["error"] = str(ex)

    def get_progress(self) -> Dict[str, Any]:
        return self.progress_state

    def stop_pipeline(self) -> bool:
        if self.pipeline:
            self.pipeline.request_stop()
            return True
        return False


# ==============================================================================
# MAIN LAUNCHER (PyWebView với Fallback Edge App Mode)
# ==============================================================================

def run_app():
    backend = SapZoneWebBackend()

    # Thử khởi chạy bằng PyWebView (Edge WebView2)
    if webview is not None:
        try:
            window = webview.create_window(
                title="SAP ZONE — Ra Đơn & In Phiếu Tự Động (v1.6)",
                html=HTML_UI,
                js_api=backend,
                width=1120,
                height=780,
                min_size=(920, 640),
                background_color="#f1f5f9"
            )
            webview.start(debug=False)
            return
        except Exception as e:
            print(f"[WEBVIEW WARN] Không thể chạy qua pywebview: {e}. Chuyển sang chế độ Local Web Server...")

    # Fallback: Chạy Local Web Server + Edge App Mode
    import http.server
    import socketserver
    import socket

    # Tìm port rảnh
    sock = socket.socket()
    sock.bind(('127.0.0.1', 0))
    port = sock.getsockname()[1]
    sock.close()

    class WebHandler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path == "/" or self.path == "/index.html":
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(HTML_UI.encode("utf-8"))
            elif self.path == "/api/status":
                data = backend.get_sap_status()
                self._send_json(data)
            elif self.path == "/api/progress":
                data = backend.get_progress()
                self._send_json(data)
            else:
                self.send_response(404)
                self.end_headers()

        def _send_json(self, obj):
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(obj).encode("utf-8"))

        def log_message(self, format, *args):
            pass

    server = socketserver.TCPServer(('127.0.0.1', port), WebHandler)
    srv_thread = threading.Thread(target=server.serve_forever, daemon=True)
    srv_thread.start()

    url = f"http://127.0.0.1:{port}"
    # Mở Edge ở App Mode (cửa sổ không viền, không tab như native app)
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ]
    edge_found = False
    for ep in edge_paths:
        if os.path.exists(ep):
            subprocess.Popen([ep, f"--app={url}", "--window-size=1120,780"])
            edge_found = True
            break
    if not edge_found:
        import webbrowser
        webbrowser.open(url)

    # Giữ tiến trình chạy
    while True:
        time.sleep(1)


if __name__ == "__main__":
    run_app()
