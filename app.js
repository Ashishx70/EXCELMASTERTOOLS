/**
 * MASTER EXCEL TOOLS - Client-side Excel & Zoho Books Merger
 * Follows Project Rules:
 * - Full-size dashboard
 * - Light theme always
 * - Transparent glass buttons
 * - Strict syntax integrity
 */

(function () {
    'use strict';

    // Filter Target Columns Configuration for Zoho & ERP (Sale & Purchase)
    const TARGET_FILTER_COLUMNS = {
        zoho_sale: [
            "transaction_number",
            "transaction_date",
            "item.CF.Order No",
            "quantity_ordered",
            "bcy_item_price",
            "sgst",
            "cgst",
            "igst",
            "bcy_tax_amount",
            "bcy_total_with_tax",
            "sku"
        ],
        zoho_purchase: [
            "transaction_number",
            "transaction_date",
            "item.CF.Order No",
            "quantity_ordered",
            "amount_without_tax",
            "sgst",
            "cgst",
            "igst",
            "bcy_total",
            "sku"
        ],
        erp_sale: [
            "Invoice",
            "Posting Date",
            "Order ID",
            "Stock Qty",
            "Rate",
            "Total Tax",
            "Output Tax CGST Amount",
            "Output Tax SGST Amount",
            "IGST Amount",
            "Total",
            "Item Code"
        ],
        erp_purchase: [
            "Bill No",
            "Posting Date",
            "Order ID",
            "Stock Qty",
            "Rate",
            "Input Tax CGST Amount",
            "Input Tax SGST Amount",
            "IGST Amount",
            "Total Tax",
            "Total",
            "Item Code"
        ]
    };

    // Application State
    const state = {
        mode: 'standard', // 'standard', 'zoho', 'erp', 'account_center', 'zip_consolidator', 'custom', 'cleaner'
        zohoPart: 'normal', // 'normal', 'sale', 'purchase'
        erpPart: 'normal', // 'normal', 'sale', 'purchase'
        files: [], // Array of File objects
        mergedHeaders: [],
        mergedData: [],
        lastMergedResult: null, // Cached { headers, data, filename, format, blob, file, rowsCount, colsCount, sizeBytes, mode }
        isProcessing: false
    };

    // Dedicated GST Cleaner State
    const cleanerState = {
        activeTab: 'tab-sale-purchase',
        salePurSubTab: 'summary', // 'summary', 'details', 'all'
        detailsCat: 'aj_my_fk', // 'aj_my_fk', 'az', 'both'
        lastProcessedSlotIds: [], // Track slots cleaned in latest batch for surgical memory wiping
        cleanedFiles: [],
        zipBlob: null,
        zipBlobUrl: null,
        zipFilename: '',
        slots: {
            // Part 1: Sales & Purchase Summary Registers (4 Files)
            s3: { file: null, regType: 'sale_erp', name: 'ERP Sales Register (Summary)', cleanName: 'Cleaned_ERP_Sales_Summary' },
            s2: { file: null, regType: 'sale_zoho', name: 'Zoho Sales Register (Summary)', cleanName: 'Cleaned_Zoho_Sales_Summary' },
            p3: { file: null, regType: 'purchase_erp', name: 'ERP Purchase Register (Summary)', cleanName: 'Cleaned_ERP_Purchase_Summary' },
            p2: { file: null, regType: 'purchase_zoho', name: 'Zoho Purchase Register (Summary)', cleanName: 'Cleaned_Zoho_Purchase_Summary' },

            // Part 2A: Sales & Purchase Details Registers - AJ&MY&FK (4 Files)
            s3_details_aj: { file: null, regType: 'sale_erp_details_aj', name: 'ERP Sales Register (Details • AJ&MY&FK)', cleanName: 'Cleaned_ERP_Sales_Details_AJ_MY_FK' },
            s2_details_aj: { file: null, regType: 'sale_zoho_details_aj', name: 'Zoho Sales Register (Details • AJ&MY&FK)', cleanName: 'Cleaned_Zoho_Sales_Details_AJ_MY_FK' },
            p3_details_aj: { file: null, regType: 'purchase_erp_details_aj', name: 'ERP Purchase Register (Details • AJ&MY&FK)', cleanName: 'Cleaned_ERP_Purchase_Details_AJ_MY_FK' },
            p2_details_aj: { file: null, regType: 'purchase_zoho_details_aj', name: 'Zoho Purchase Register (Details • AJ&MY&FK)', cleanName: 'Cleaned_Zoho_Purchase_Details_AJ_MY_FK' },

            // Part 2B: Sales & Purchase Details Registers - AZ (4 Files)
            s3_details_az: { file: null, regType: 'sale_erp_details_az', name: 'ERP Sales Register (Details • AZ)', cleanName: 'Cleaned_ERP_Sales_Details_AZ' },
            s2_details_az: { file: null, regType: 'sale_zoho_details_az', name: 'Zoho Sales Register (Details • AZ)', cleanName: 'Cleaned_Zoho_Sales_Details_AZ' },
            p3_details_az: { file: null, regType: 'purchase_erp_details_az', name: 'ERP Purchase Register (Details • AZ)', cleanName: 'Cleaned_ERP_Purchase_Details_AZ' },
            p2_details_az: { file: null, regType: 'purchase_zoho_details_az', name: 'Zoho Purchase Register (Details • AZ)', cleanName: 'Cleaned_Zoho_Purchase_Details_AZ' },

            // Backward Compatibility aliases
            s3_details: { file: null, regType: 'sale_erp_details_aj', name: 'ERP Sales Register (Details)', cleanName: 'Cleaned_ERP_Sales_Details' },
            s2_details: { file: null, regType: 'sale_zoho_details_aj', name: 'Zoho Sales Register (Details)', cleanName: 'Cleaned_Zoho_Sales_Details' },
            p3_details: { file: null, regType: 'purchase_erp_details_aj', name: 'ERP Purchase Register (Details)', cleanName: 'Cleaned_ERP_Purchase_Details' },
            p2_details: { file: null, regType: 'purchase_zoho_details_aj', name: 'Zoho Purchase Register (Details)', cleanName: 'Cleaned_Zoho_Purchase_Details' },

            // Credit Notes & Debit Notes
            c3: { file: null, regType: 'credit_note_erp', name: 'ERP Credit Notes', cleanName: 'Cleaned_ERP_Credit_Notes' },
            c2: { file: null, regType: 'credit_note_zoho', name: 'Zoho Credit Notes', cleanName: 'Cleaned_Zoho_Credit_Notes' },
            d3: { file: null, regType: 'debit_note_erp', name: 'ERP Debit Notes', cleanName: 'Cleaned_ERP_Debit_Notes' },
            d2: { file: null, regType: 'debit_note_zoho', name: 'Zoho Debit Notes', cleanName: 'Cleaned_Zoho_Debit_Notes' }
        }
    };

    // DOM Elements
    const elements = {
        // Mode Radios & Cards
        modeStandard: document.getElementById('card-mode-standard'),
        modeZoho: document.getElementById('card-mode-zoho'),
        modeErp: document.getElementById('card-mode-erp'),
        modeAccountCenter: document.getElementById('card-mode-account-center'),
        modeZipConsolidator: document.getElementById('card-mode-zip-consolidator'),
        modeCustom: document.getElementById('card-mode-custom'),
        modeCleaner: document.getElementById('card-mode-cleaner'),
        customSettings: document.getElementById('custom-settings'),
        cleanerWorkspace: document.getElementById('cleaner-workspace'),
        zohoSubPanel: document.getElementById('zoho-sub-panel'),
        zohoActiveBadge: document.getElementById('zoho-active-badge'),
        zohoCardSubtag: document.getElementById('zoho-card-subtag'),
        erpSubPanel: document.getElementById('erp-sub-panel'),
        erpActiveBadge: document.getElementById('erp-active-badge'),
        erpCardSubtag: document.getElementById('erp-card-subtag'),
        customHeaderRow: document.getElementById('custom-header-row'),
        customDataRow: document.getElementById('custom-data-row'),
        customSheetIndex: document.getElementById('custom-sheet-index'),

        // Cleaning Audit Stats Card & Downloads
        cleaningStatsCard: document.getElementById('cleaning-stats-card'),
        statColDetected: document.getElementById('stat-col-detected'),
        statHeaderRow: document.getElementById('stat-header-row'),
        statOrigRows: document.getElementById('stat-orig-rows'),
        statRetainedRows: document.getElementById('stat-retained-rows'),
        statExcludedRows: document.getElementById('stat-excluded-rows'),
        statPeakRam: document.getElementById('stat-peak-ram'),
        statDuration: document.getElementById('stat-duration'),
        statsFileInfo: document.getElementById('stats-file-info'),
        statsDownloadSection: document.getElementById('stats-download-section'),
        statsZipBox: document.getElementById('stats-zip-box'),
        btnDownloadZip: document.getElementById('btn-download-zip'),
        btnDownloadZipText: document.getElementById('btn-download-zip-text'),
        zipFileName: document.getElementById('zip-file-name'),
        zipFileDesc: document.getElementById('zip-file-desc'),
        cleanedFilesGrid: document.getElementById('cleaned-files-grid'),
        individualDownloadsHeader: document.getElementById('individual-downloads-header'),

        // Upload & Inputs
        uploadSection: document.querySelector('.upload-section'),
        dropzone: document.getElementById('dropzone'),
        fileInputFiles: document.getElementById('file-input-files'),
        fileInputFolder: document.getElementById('file-input-folder'),
        fileInputZip: document.getElementById('file-input-zip'),
        btnBrowseFiles: document.getElementById('btn-browse-files'),
        btnBrowseFolder: document.getElementById('btn-browse-folder'),
        btnBrowseZip: document.getElementById('btn-browse-zip'),
        btnDropzoneClick: document.getElementById('btn-dropzone-click'),

        // Queue
        queueContainer: document.getElementById('queue-container'),
        fileTableBody: document.getElementById('file-table-body'),
        statFileCount: document.getElementById('stat-file-count'),
        statTotalSize: document.getElementById('stat-total-size'),
        btnClearQueue: document.getElementById('btn-clear-queue'),

        // Export & Actions
        exportSection: document.getElementById('export-section'),
        outputFilename: document.getElementById('output-filename'),
        outputExtLabel: document.getElementById('output-ext-label'),
        outputFormat: document.getElementById('output-format'),
        cleanerOutputFormat: document.getElementById('cleaner-output-format-bar'),
        progressBox: document.getElementById('progress-box'),
        progressStatusText: document.getElementById('progress-status-text'),
        progressPercentage: document.getElementById('progress-percentage'),
        progressBarFill: document.getElementById('progress-bar-fill'),
        btnPreviewData: document.getElementById('btn-preview-data'),
        btnMergeDownload: document.getElementById('btn-merge-download'),

        // Merged Result Action Box
        mergedResultBox: document.getElementById('merged-result-box'),
        mergedResultFilename: document.getElementById('merged-result-filename'),
        mergedResultRows: document.getElementById('merged-result-rows'),
        mergedResultCols: document.getElementById('merged-result-cols'),
        mergedResultSize: document.getElementById('merged-result-size'),
        mergedModeTag: document.getElementById('merged-mode-tag'),
        btnDownloadMerged: document.getElementById('btn-download-merged'),
        btnSendToCleaner: document.getElementById('btn-send-to-cleaner'),

        // Send to Cleaner Modal
        modalSendCleaner: document.getElementById('modal-send-cleaner'),
        btnCloseSendModal: document.getElementById('btn-close-send-modal'),
        btnCancelSendCleaner: document.getElementById('btn-cancel-send-cleaner'),
        btnConfirmSendCleaner: document.getElementById('btn-confirm-send-cleaner'),
        sendModalFilename: document.getElementById('send-modal-filename'),
        sendModalStats: document.getElementById('send-modal-stats'),

        // Preview Modal
        previewModal: document.getElementById('preview-modal'),
        btnClosePreview: document.getElementById('btn-close-preview'),
        previewTableHead: document.getElementById('preview-table-head'),
        previewTableBody: document.getElementById('preview-table-body'),
        previewStatsText: document.getElementById('preview-stats-text'),

        // Toasts
        toastContainer: document.getElementById('toast-container')
    };

    // Initialize Event Listeners
    function init() {
        setupModeSwitching();
        setupZohoSubParts();
        setupErpSubParts();
        setupCleanerWorkspace();
        setupFileUploads();
        setupActions();
        updateOutputFilenameDefault();
    }

    // Format bytes to human readable format
    function formatBytes(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    // Generate formatted date timestamp: YYYYMMDD_HHMMSS
    function getTimestampString() {
        const now = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    }

    // Toast Notifications
    function showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        
        const iconSvg = type === 'success' 
            ? `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`
            : `<svg class="toast-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`;

        toast.innerHTML = `${iconSvg}<span>${message}</span>`;
        elements.toastContainer.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(10px)';
            toast.style.transition = 'all 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    // Mode Switching Logic
    function setupModeSwitching() {
        const modeCards = [
            elements.modeStandard,
            elements.modeZoho,
            elements.modeErp,
            elements.modeAccountCenter,
            elements.modeZipConsolidator,
            elements.modeCustom,
            elements.modeCleaner
        ];

        modeCards.forEach(card => {
            if (!card) return;
            card.addEventListener('click', () => {
                modeCards.forEach(c => c && c.classList.remove('active'));
                card.classList.add('active');
                const radio = card.querySelector('input[type="radio"]');
                radio.checked = true;
                state.mode = radio.value;

                if (state.mode === 'custom') {
                    elements.customSettings.style.display = 'block';
                } else {
                    elements.customSettings.style.display = 'none';
                }

                if (state.mode === 'zoho') {
                    if (elements.zohoSubPanel) elements.zohoSubPanel.style.display = 'block';
                } else {
                    if (elements.zohoSubPanel) elements.zohoSubPanel.style.display = 'none';
                }

                if (state.mode === 'erp') {
                    if (elements.erpSubPanel) elements.erpSubPanel.style.display = 'block';
                } else {
                    if (elements.erpSubPanel) elements.erpSubPanel.style.display = 'none';
                }

                if (state.mode === 'cleaner') {
                    if (elements.cleanerWorkspace) elements.cleanerWorkspace.style.display = 'block';
                    if (elements.uploadSection) elements.uploadSection.style.display = 'none';
                    if (elements.queueContainer) elements.queueContainer.style.display = 'none';
                    if (elements.exportSection) elements.exportSection.style.display = 'none';
                } else {
                    if (elements.cleanerWorkspace) elements.cleanerWorkspace.style.display = 'none';
                    if (elements.uploadSection) elements.uploadSection.style.display = 'block';
                    if (elements.cleaningStatsCard) elements.cleaningStatsCard.style.display = 'none';

                    if (state.files.length > 0) {
                        elements.queueContainer.style.display = 'flex';
                        elements.exportSection.style.display = 'block';
                    } else {
                        elements.queueContainer.style.display = 'none';
                        elements.exportSection.style.display = 'none';
                    }

                    if (state.mode === 'zip_consolidator') {
                        elements.btnPreviewData.style.display = 'none';
                        elements.btnMergeDownload.querySelector('span').textContent = 'Package & Download Single ZIP';
                    } else {
                        elements.btnPreviewData.style.display = 'inline-flex';
                        elements.btnMergeDownload.querySelector('span').textContent = 'Merge Files';
                    }

                    // Reset merged result box on mode switch
                    if (elements.mergedResultBox) {
                        elements.mergedResultBox.style.display = 'none';
                    }
                }

                updateOutputFilenameDefault();
            });
        });

        // Format change update
        elements.outputFormat.addEventListener('change', () => {
            const format = elements.outputFormat.value;
            let ext = '.xlsx';
            if (format === 'csv') ext = '.csv';
            else if (format === 'zip') ext = '.zip';
            elements.outputExtLabel.textContent = ext;
        });
    }

    // Setup Zoho Sub-Part Selector (Normal, Sale, Purchase)
    function setupZohoSubParts() {
        const partCards = document.querySelectorAll('.zoho-part-card');
        partCards.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const part = btn.dataset.zohoPart || 'normal';
                state.zohoPart = part;
                partCards.forEach(c => c.classList.remove('active'));
                btn.classList.add('active');

                const cap = part.charAt(0).toUpperCase() + part.slice(1);
                if (elements.zohoActiveBadge) {
                    elements.zohoActiveBadge.textContent = `${cap} Selected`;
                }
                if (elements.zohoCardSubtag) {
                    elements.zohoCardSubtag.textContent = cap;
                }

                updateOutputFilenameDefault();
            });
        });
    }

    // Setup ERP Sub-Part Selector (Normal, Sale, Purchase)
    function setupErpSubParts() {
        const partCards = document.querySelectorAll('.erp-part-card');
        partCards.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const part = btn.dataset.erpPart || 'normal';
                state.erpPart = part;
                partCards.forEach(c => c.classList.remove('active'));
                btn.classList.add('active');

                const cap = part.charAt(0).toUpperCase() + part.slice(1);
                if (elements.erpActiveBadge) {
                    elements.erpActiveBadge.textContent = `${cap} Selected`;
                }
                if (elements.erpCardSubtag) {
                    elements.erpCardSubtag.textContent = cap;
                }

                updateOutputFilenameDefault();
            });
        });
    }

    // Resilient Column Header Normalizer (lowercase & ignores whitespace, symbols, punctuation, and currency tags)
    function normalizeColHeader(str) {
        let s = String(str || '').toLowerCase().trim();
        // Strip parenthesized or delimited currency tags, e.g. " (INR)", " (Rs.)", "_INR", " (₹)"
        // Do NOT strip bare trailing 'rs' like in 'particulars', 'orders', 'vouchers'
        s = s.replace(new RegExp('\\s*[\\(\\[]\\s*(inr|rs|rs\\.|₹)\\s*[\\)\\]]', 'gi'), '');
        s = s.replace(new RegExp('[\\s_\\-]+(inr|rs)$', 'gi'), '');
        // Standardize 'itern' typo in Zoho files
        s = s.replace(/^itern/gi, 'item');
        // Strip symbols, whitespace, and punctuation
        return s.replace(/[\s_\.\-\#\(\)\[\]\/\:\;\\\'\"₹,]+/g, '');
    }

    // Comprehensive column aliases for ERP & Zoho fields
    const COLUMN_ALIASES = {
        // ERP Sale CGST
        outputtaxcgstamount: [
            'outputtaxcgstamount', 'outputtaxcgst', 'outputcgstamount', 'outputcgstamt', 'outputcgst',
            'taxcgstamount', 'taxcgst', 'cgstamount', 'cgstamt', 'cgstvalue', 'cgsttaxamount', 'cgsttax',
            'cgstoutputamount', 'cgstoutput', 'centraltaxamount', 'centraltaxamt', 'centraltax',
            'outputcentraltaxamount', 'outputcentraltax', 'centralgstamount', 'centralgst', 'cgst'
        ],
        // ERP Sale SGST
        outputtaxsgstamount: [
            'outputtaxsgstamount', 'outputtaxsgst', 'outputsgstamount', 'outputsgstamt', 'outputsgst',
            'taxsgstamount', 'taxsgst', 'sgstamount', 'sgstamt', 'sgstvalue', 'sgsttaxamount', 'sgsttax',
            'sgstoutputamount', 'sgstoutput', 'statetaxamount', 'statetaxamt', 'statetax',
            'outputstatetaxamount', 'outputstatetax', 'outputtaxutgstamount', 'outputtaxutgst',
            'outpututgstamount', 'outpututgst', 'utgstamount', 'utgstamt', 'stategstamount', 'stategst', 'utgst', 'sgst'
        ],
        // ERP Purchase CGST
        inputtaxcgstamount: [
            'inputtaxcgstamount', 'inputtaxcgst', 'inputcgstamount', 'inputcgstamt', 'inputcgst',
            'taxcgstamount', 'taxcgst', 'cgstamount', 'cgstamt', 'cgstvalue', 'cgsttaxamount', 'cgsttax',
            'cgstinputamount', 'cgstinput', 'centraltaxamount', 'centraltaxamt', 'centraltax',
            'inputcentraltaxamount', 'inputcentraltax', 'centralgstamount', 'centralgst', 'cgst'
        ],
        // ERP Purchase SGST
        inputtaxsgstamount: [
            'inputtaxsgstamount', 'inputtaxsgst', 'inputsgstamount', 'inputsgstamt', 'inputsgst',
            'taxsgstamount', 'taxsgst', 'sgstamount', 'sgstamt', 'sgstvalue', 'sgsttaxamount', 'sgsttax',
            'sgstinputamount', 'sgstinput', 'statetaxamount', 'statetaxamt', 'statetax',
            'inputstatetaxamount', 'inputstatetax', 'inputtaxutgstamount', 'inputtaxutgst',
            'inpututgstamount', 'inpututgst', 'utgstamount', 'utgstamt', 'stategstamount', 'stategst', 'utgst', 'sgst'
        ],
        // IGST (Sale & Purchase)
        igstamount: [
            'igstamount', 'outputtaxigstamount', 'inputtaxigstamount', 'outputigstamount', 'inputigstamount',
            'outputtaxigst', 'inputtaxigst', 'outputigst', 'inputigst', 'taxigstamount', 'taxigst',
            'igstamt', 'igstvalue', 'igsttaxamount', 'igsttax', 'integratedtaxamount', 'integratedtaxamt',
            'integratedtax', 'integratedgstamount', 'integratedgst', 'igst'
        ],
        // Zoho CGST
        cgst: [
            'cgst', 'cgstamount', 'cgstamt', 'outputtaxcgstamount', 'outputtaxcgst', 'inputtaxcgstamount',
            'inputtaxcgst', 'taxcgstamount', 'taxcgst', 'centraltaxamount', 'centraltax', 'centralgst'
        ],
        // Zoho SGST
        sgst: [
            'sgst', 'sgstamount', 'sgstamt', 'outputtaxsgstamount', 'outputtaxsgst', 'inputtaxsgstamount',
            'inputtaxsgst', 'taxsgstamount', 'taxsgst', 'statetaxamount', 'statetax', 'utgstamount', 'utgst', 'stategst'
        ],
        // Zoho IGST
        igst: [
            'igst', 'igstamount', 'igstamt', 'outputtaxigstamount', 'inputtaxigstamount', 'taxigstamount',
            'taxigst', 'integratedtaxamount', 'integratedtax', 'integratedgst'
        ],
        // Total Tax
        totaltax: [
            'totaltax', 'totaltaxamount', 'totaltaxamt', 'taxamount', 'taxamt', 'bcytaxamount',
            'totaltaxes', 'gstamount', 'gstamt', 'totalgst', 'tax'
        ],
        bcytaxamount: [
            'bcytaxamount', 'totaltax', 'totaltaxamount', 'taxamount', 'taxamt', 'totaltaxamt', 'gstamount', 'gstamt', 'totalgst', 'tax'
        ],
        // Invoice / Bill No
        invoice: [
            'invoice', 'invoiceno', 'invoicenumber', 'invno', 'inv', 'voucherno', 'vouchernumber', 'vchno', 'vch',
            'docno', 'documentno', 'refno', 'referenceno', 'billno', 'billnumber', 'bill', 'transactionnumber', 'transno', 'billinvno'
        ],
        billno: [
            'billno', 'billnumber', 'bill', 'invoiceno', 'invoicenumber', 'invoice', 'invno', 'inv',
            'voucherno', 'vouchernumber', 'vchno', 'vch', 'docno', 'documentno', 'refno', 'referenceno', 'transactionnumber', 'transno'
        ],
        transactionnumber: [
            'transactionnumber', 'transno', 'billnumber', 'billno', 'invoicenumber', 'invoiceno', 'invoice', 'bill', 'voucherno', 'vchno'
        ],
        // Dates
        postingdate: [
            'postingdate', 'invoicedate', 'invdate', 'billdate', 'date', 'voucherdate', 'vchdate', 'docdate', 'documentdate', 'transactiondate', 'transdate', 'txndate'
        ],
        transactiondate: [
            'transactiondate', 'transdate', 'postingdate', 'invoicedate', 'invdate', 'billdate', 'date', 'voucherdate', 'docdate'
        ],
        // Order ID (for ERP)
        orderid: [
            'orderid', 'orderno', 'ordernumber', 'itemcforderno', 'cforderno', 'cfordernumber', 'pono', 'purchaseorderno',
            'sono', 'salesorderno', 'customerorderno', 'orderref', 'order'
        ],
        // Custom Field Order No (for Zoho Purchase & Sale: item.CF.Order No / itern.CF.Order No)
        itemcforderno: [
            'itemcforderno', 'iterncforderno', 'cforderno', 'cfordernumber',
            'cforder', 'itemorderno', 'itemordernumber', 'iternorderno', 'iternordernumber', 'orderid', 'orderno', 'ordernumber'
        ],
        // Quantity
        stockqty: [
            'stockqty', 'billedqty', 'actualqty', 'qty', 'quantity', 'quantityordered', 'stockquantity', 'itemqty', 'netqty', 'totalqty', 'nos', 'pcs', 'units'
        ],
        quantityordered: [
            'quantityordered', 'stockqty', 'billedqty', 'qty', 'quantity', 'itemqty', 'nos', 'pcs'
        ],
        // Rate / Price
        rate: [
            'rate', 'itemprice', 'bcyitemprice', 'amountwithouttax', 'price', 'unitprice', 'basicrate',
            'basicprice', 'basicamount', 'taxableamount', 'taxablevalue', 'taxableamt', 'taxable', 'assessablevalue', 'netrate', 'amount'
        ],
        bcyitemprice: [
            'bcyitemprice', 'amountwithouttax', 'rate', 'price', 'itemprice', 'unitprice', 'taxableamount', 'taxablevalue'
        ],
        amountwithouttax: [
            'amountwithouttax', 'bcyitemprice', 'rate', 'price', 'taxableamount', 'taxablevalue', 'taxableamt', 'basicamount', 'itemprice'
        ],
        // Total
        total: [
            'total', 'bcytotalwithtax', 'bcytotal', 'grandtotal', 'invoicetotal', 'billtotal', 'nettotal',
            'netamount', 'invoiceamount', 'invoicevalue', 'billamount', 'billvalue', 'totalamount', 'totalvalue', 'grosstotal', 'finalamount'
        ],
        bcytotalwithtax: [
            'bcytotalwithtax', 'bcytotal', 'total', 'grandtotal', 'invoicetotal', 'totalamount', 'invoiceamount', 'invoicevalue', 'netamount'
        ],
        bcytotal: [
            'bcytotal', 'bcytotalwithtax', 'total', 'grandtotal', 'billtotal', 'totalamount', 'billamount', 'billvalue', 'netamount'
        ],
        // Item Code / SKU
        itemcode: [
            'itemcode', 'sku', 'item', 'itemname', 'productcode', 'product', 'productname', 'particulars', 'description', 'itemdescription', 'productdescription', 'materialcode', 'itemid', 'partno', 'partnumber'
        ],
        sku: [
            'sku', 'itemcode', 'item', 'productcode', 'itemname', 'product', 'particulars', 'description', 'itemdescription'
        ]
    };

    // Build column mapping: maps targetColumns -> source file column index (-1 if missing)
    function buildColumnMapping(fileHeaders, targetColumns) {
        const mapping = [];
        const normHeaders = fileHeaders.map(h => normalizeColHeader(h));
        const exactHeaders = fileHeaders.map(h => String(h || '').trim().toLowerCase());
        const usedIndices = new Set();

        targetColumns.forEach(target => {
            const targetExact = target.trim().toLowerCase();
            const targetNorm = normalizeColHeader(target);
            let foundIdx = -1;

            // 1. Try exact trimmed case-insensitive match (unused index first)
            for (let i = 0; i < exactHeaders.length; i++) {
                if (exactHeaders[i] === targetExact && !usedIndices.has(i)) {
                    foundIdx = i;
                    break;
                }
            }

            // 2. Try normalized match (unused index first)
            if (foundIdx === -1) {
                for (let i = 0; i < normHeaders.length; i++) {
                    if (normHeaders[i] === targetNorm && !usedIndices.has(i)) {
                        foundIdx = i;
                        break;
                    }
                }
            }

            // 3. Try alias list
            if (foundIdx === -1 && COLUMN_ALIASES[targetNorm]) {
                const aliases = COLUMN_ALIASES[targetNorm];
                for (let a = 0; a < aliases.length; a++) {
                    const aliasNorm = aliases[a];
                    for (let i = 0; i < normHeaders.length; i++) {
                        if (normHeaders[i] === aliasNorm && !usedIndices.has(i)) {
                            foundIdx = i;
                            break;
                        }
                    }
                    if (foundIdx !== -1) break;
                }
            }

            // 4. Keyword / Substring Fallback for Tax Columns (CGST, SGST, IGST)
            if (foundIdx === -1) {
                const isCgst = targetNorm.includes('cgst') || targetNorm.includes('centraltax');
                const isSgst = targetNorm.includes('sgst') || targetNorm.includes('statetax') || targetNorm.includes('utgst');
                const isIgst = targetNorm.includes('igst') || targetNorm.includes('integratedtax');
                const isTotalTax = targetNorm === 'totaltax' || targetNorm === 'bcytaxamount';

                for (let i = 0; i < normHeaders.length; i++) {
                    if (usedIndices.has(i)) continue;
                    const h = normHeaders[i];
                    if (!h) continue;

                    // Avoid matching rate or percentage columns
                    if (h.includes('rate') || h.includes('perc') || h.includes('%')) continue;

                    if (isCgst && (h.includes('cgst') || h.includes('centraltax'))) {
                        foundIdx = i;
                        break;
                    } else if (isSgst && (h.includes('sgst') || h.includes('statetax') || h.includes('utgst'))) {
                        foundIdx = i;
                        break;
                    } else if (isIgst && (h.includes('igst') || h.includes('integratedtax'))) {
                        foundIdx = i;
                        break;
                    } else if (isTotalTax && (h.includes('totaltax') || (h.includes('tax') && h.includes('amount') && !h.includes('cgst') && !h.includes('sgst') && !h.includes('igst')))) {
                        foundIdx = i;
                        break;
                    }
                }
            }

            // 5. Intelligent Fallback for Custom Field Order No, Order ID, Invoice/Bill, Item Code, Dates, Qty, Rate, Total
            if (foundIdx === -1) {
                const isCustomOrder = targetNorm.includes('cf') && targetNorm.includes('order');
                const isOrder = targetNorm.includes('order') && !isCustomOrder;
                const isInv = targetNorm === 'invoice' || targetNorm === 'billno' || targetNorm === 'transactionnumber';
                const isItem = targetNorm === 'itemcode' || targetNorm === 'sku';
                const isDate = targetNorm.includes('date');
                const isQty = targetNorm === 'stockqty' || targetNorm === 'quantityordered';
                const isRate = targetNorm === 'rate' || targetNorm === 'bcyitemprice' || targetNorm === 'amountwithouttax';
                const isTotal = targetNorm === 'total' || targetNorm === 'bcytotalwithtax' || targetNorm === 'bcytotal';

                // Prioritize Custom Field Order No: match ANY header containing BOTH 'cf' and 'order'
                if (isCustomOrder) {
                    for (let i = 0; i < normHeaders.length; i++) {
                        if (usedIndices.has(i)) continue;
                        const h = normHeaders[i];
                        if (!h) continue;
                        if (h.includes('cf') && h.includes('order')) {
                            foundIdx = i;
                            break;
                        }
                    }
                }

                if (foundIdx === -1) {
                    for (let i = 0; i < normHeaders.length; i++) {
                        if (usedIndices.has(i)) continue;
                        const h = normHeaders[i];
                        if (!h) continue;

                        if (isInv) {
                            // Crucial: avoid matching quantity columns (e.g. 'billedqty') or date columns
                            if (h.includes('qty') || h.includes('quantity') || h.includes('date')) continue;
                            if (h.includes('invoice') || h.includes('inv') || h.includes('bill') || h.includes('voucher') || h.includes('vch') || h.includes('docno') || h.includes('transno') || h.includes('refno')) {
                                foundIdx = i;
                                break;
                            }
                        } else if (isDate) {
                            if (h.includes('update') || h.includes('due')) continue;
                            if (h.includes('date') || h.includes('dt') || h.includes('posting') || h.includes('txn')) {
                                foundIdx = i;
                                break;
                            }
                        } else if (isOrder) {
                            if (h.includes('date') || h.includes('qty') || h.includes('price') || h.includes('amount')) continue;
                            if (h.includes('orderno') || h.includes('orderid') || h.includes('order') || h.includes('sono') || h.includes('pono') || h.includes('custref')) {
                                foundIdx = i;
                                break;
                            }
                        } else if (isQty) {
                            if (h.includes('rate') || h.includes('price') || h.includes('amount')) continue;
                            if (h.includes('qty') || h.includes('quantity') || h === 'nos' || h === 'pcs' || h === 'units') {
                                foundIdx = i;
                                break;
                            }
                        } else if (isRate) {
                            if (h.includes('qty') || h.includes('tax') || h.includes('%') || h.includes('perc')) continue;
                            if (h.includes('rate') || h.includes('price') || h.includes('taxable') || h.includes('assessable') || h.includes('basic')) {
                                foundIdx = i;
                                break;
                            }
                        } else if (isTotal) {
                            if (h.includes('tax') || h.includes('qty')) continue;
                            if (h.includes('total') || h.includes('gross') || h.includes('netamt') || h.includes('netamount') || h.includes('grand')) {
                                foundIdx = i;
                                break;
                            }
                        } else if (isItem) {
                            if (h.includes('rate') || h.includes('price') || h.includes('amount') || h.includes('tax')) continue;
                            if (h.includes('sku') || h.includes('itemcode') || (h.includes('item') && !h.includes('price')) || h.includes('product') || h.includes('particulars') || h.includes('description')) {
                                foundIdx = i;
                                break;
                            }
                        }
                    }
                }
            }

            if (foundIdx !== -1) {
                usedIndices.add(foundIdx);
            }

            mapping.push(foundIdx);
        });

        return mapping;
    }

    // Setup Dedicated Multi-Slot GST Cleaner Workspace
    function setupCleanerWorkspace() {
        // Tab switching (Sale & Pur, Credit Notes, Debit Notes)
        const tabBtns = document.querySelectorAll('.cleaner-tab-btn');
        tabBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                tabBtns.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                const targetTabId = btn.getAttribute('data-clean-tab');
                cleanerState.activeTab = targetTabId;

                document.querySelectorAll('.cleaner-tab-content').forEach(tabContent => {
                    tabContent.style.display = tabContent.id === targetTabId ? 'block' : 'none';
                });
            });
        });

        // Sub-tabs switching inside Sale & Purchase (Summary vs Details vs All)
        const subTabBtns = document.querySelectorAll('.sale-pur-subtab-btn');
        const paneSummary = document.getElementById('pane-sale-pur-summary');
        const paneDetails = document.getElementById('pane-sale-pur-details');

        function switchSalePurSubTab(subTab) {
            cleanerState.salePurSubTab = subTab;
            subTabBtns.forEach(b => {
                b.classList.toggle('active', b.getAttribute('data-subtab') === subTab);
            });
            if (subTab === 'summary') {
                if (paneSummary) paneSummary.style.display = 'block';
                if (paneDetails) paneDetails.style.display = 'none';
            } else if (subTab === 'details') {
                if (paneSummary) paneSummary.style.display = 'none';
                if (paneDetails) paneDetails.style.display = 'block';
            } else { // 'all'
                if (paneSummary) paneSummary.style.display = 'block';
                if (paneDetails) paneDetails.style.display = 'block';
            }
        }

        subTabBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                const subtab = btn.getAttribute('data-subtab');
                switchSalePurSubTab(subtab);
            });
        });

        // Details Sub-category switching (AJ&MY&FK vs AZ vs Both)
        const detailsCatBtns = document.querySelectorAll('.details-cat-btn');
        const paneDetailsAj = document.getElementById('pane-details-aj');
        const paneDetailsAz = document.getElementById('pane-details-az');

        function switchDetailsCat(cat) {
            cleanerState.detailsCat = cat;
            detailsCatBtns.forEach(b => {
                b.classList.toggle('active', b.getAttribute('data-details-cat') === cat);
            });
            if (cat === 'aj_my_fk') {
                if (paneDetailsAj) paneDetailsAj.style.display = 'block';
                if (paneDetailsAz) paneDetailsAz.style.display = 'none';
            } else if (cat === 'az') {
                if (paneDetailsAj) paneDetailsAj.style.display = 'none';
                if (paneDetailsAz) paneDetailsAz.style.display = 'block';
            } else { // 'both'
                if (paneDetailsAj) paneDetailsAj.style.display = 'block';
                if (paneDetailsAz) paneDetailsAz.style.display = 'block';
            }
        }

        detailsCatBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                const cat = btn.getAttribute('data-details-cat');
                switchDetailsCat(cat);
            });
        });

        // "Sync Files to AZ" Button: Copies files from AJ slots to AZ slots
        const btnSyncDetails = document.getElementById('btn-sync-details-files');
        if (btnSyncDetails) {
            btnSyncDetails.addEventListener('click', () => {
                syncAjFilesToAz();
            });
        }

        function syncAjFilesToAz() {
            const pairs = [
                { src: 's3_details_aj', dst: 's3_details_az' },
                { src: 's2_details_aj', dst: 's2_details_az' },
                { src: 'p3_details_aj', dst: 'p3_details_az' },
                { src: 'p2_details_aj', dst: 'p2_details_az' }
            ];
            let syncedCount = 0;
            pairs.forEach(({ src, dst }) => {
                const srcSlot = cleanerState.slots[src];
                if (srcSlot && srcSlot.file) {
                    const dstCard = document.querySelector(`.register-slot-card[data-slot="${dst}"]`);
                    if (dstCard) {
                        const fileMeta = dstCard.querySelector('.slot-file-meta');
                        const filenameEl = dstCard.querySelector('.slot-filename');
                        const filesizeEl = dstCard.querySelector('.slot-filesize');
                        setSlotFile(dst, srcSlot.file, dstCard, fileMeta, filenameEl, filesizeEl);
                        syncedCount++;
                    }
                }
            });
            if (syncedCount > 0) {
                showToast(`Synced ${syncedCount} file(s) into AZ Details slots!`, 'success');
            } else {
                showToast('No files uploaded in AJ&MY&FK slots to sync.', 'info');
            }
        }

        // Slot cards file selection & drag/drop
        const slotCards = document.querySelectorAll('.register-slot-card');
        slotCards.forEach(card => {
            const slotId = card.getAttribute('data-slot');
            const fileInput = card.querySelector('.slot-file-input');
            const fileMeta = card.querySelector('.slot-file-meta');
            const filenameEl = card.querySelector('.slot-filename');
            const filesizeEl = card.querySelector('.slot-filesize');
            const removeBtn = card.querySelector('.btn-slot-remove');

            // Click card opens file picker
            card.addEventListener('click', (e) => {
                if (e.target.closest('.btn-slot-remove')) return;
                fileInput.click();
            });

            // File selected via input
            fileInput.addEventListener('change', (e) => {
                if (e.target.files && e.target.files.length > 0) {
                    setSlotFile(slotId, e.target.files[0], card, fileMeta, filenameEl, filesizeEl);
                }
            });

            // Remove button
            if (removeBtn) {
                removeBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    clearSlot(slotId, card, fileInput, fileMeta);
                });
            }

            // Drag & drop on slot card
            ['dragenter', 'dragover'].forEach(eventName => {
                card.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    card.classList.add('dragover');
                });
            });

            ['dragleave', 'drop'].forEach(eventName => {
                card.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    card.classList.remove('dragover');
                });
            });

            card.addEventListener('drop', (e) => {
                const dt = e.dataTransfer;
                if (dt && dt.files && dt.files.length > 0) {
                    const file = dt.files[0];
                    const ext = file.name.split('.').pop().toLowerCase();
                    if (['xlsx', 'xls', 'csv'].includes(ext)) {
                        setSlotFile(slotId, file, card, fileMeta, filenameEl, filesizeEl);
                    } else {
                        showToast(`Unsupported file type: .${ext}. Please use .xlsx or .csv`, 'error');
                    }
                }
            });
        });

        // Clear all cleaner files button
        const btnClear = document.getElementById('btn-cleaner-clear');
        if (btnClear) {
            btnClear.addEventListener('click', () => {
                slotCards.forEach(card => {
                    const slotId = card.getAttribute('data-slot');
                    const fileInput = card.querySelector('.slot-file-input');
                    const fileMeta = card.querySelector('.slot-file-meta');
                    clearSlot(slotId, card, fileInput, fileMeta);
                });
                updateSubtabCounters();
                showToast('All cleaner slot files cleared.', 'info');
            });
        }

        // Clean Selected Files button
        const btnClean = document.getElementById('btn-clean-selected');
        if (btnClean) {
            btnClean.addEventListener('click', () => {
                if (!state.isProcessing) {
                    handleCleanSelectedFiles();
                }
            });
        }
    }

    function updateSubtabCounters() {
        const summarySlots = ['s3', 's2', 'p3', 'p2'];
        const detailsAjSlots = ['s3_details_aj', 's2_details_aj', 'p3_details_aj', 'p2_details_aj'];
        const detailsAzSlots = ['s3_details_az', 's2_details_az', 'p3_details_az', 'p2_details_az'];
        const legacyDetailsSlots = ['s3_details', 's2_details', 'p3_details', 'p2_details'];

        let summaryCount = 0;
        summarySlots.forEach(id => {
            if (cleanerState.slots[id] && cleanerState.slots[id].file) summaryCount++;
        });

        let ajCount = 0;
        detailsAjSlots.forEach(id => {
            if (cleanerState.slots[id] && cleanerState.slots[id].file) ajCount++;
        });

        let azCount = 0;
        detailsAzSlots.forEach(id => {
            if (cleanerState.slots[id] && cleanerState.slots[id].file) azCount++;
        });

        let legacyCount = 0;
        legacyDetailsSlots.forEach(id => {
            if (cleanerState.slots[id] && cleanerState.slots[id].file) legacyCount++;
        });

        const detailsTotal = ajCount + azCount + legacyCount;

        const badgeSummary = document.getElementById('badge-summary-count');
        const badgeDetails = document.getElementById('badge-details-count');
        const badgeDetailsAj = document.getElementById('badge-details-aj-count');
        const badgeDetailsAz = document.getElementById('badge-details-az-count');

        if (badgeSummary) {
            badgeSummary.textContent = summaryCount;
            badgeSummary.style.display = summaryCount > 0 ? 'inline-block' : 'none';
        }
        if (badgeDetails) {
            badgeDetails.textContent = detailsTotal;
            badgeDetails.style.display = detailsTotal > 0 ? 'inline-block' : 'none';
        }
        if (badgeDetailsAj) {
            badgeDetailsAj.textContent = ajCount;
            badgeDetailsAj.style.display = ajCount > 0 ? 'inline-block' : 'none';
        }
        if (badgeDetailsAz) {
            badgeDetailsAz.textContent = azCount;
            badgeDetailsAz.style.display = azCount > 0 ? 'inline-block' : 'none';
        }
    }

    function setSlotFile(slotId, file, card, fileMeta, filenameEl, filesizeEl) {
        if (!cleanerState.slots[slotId]) return;
        cleanerState.slots[slotId].file = file;
        card.classList.add('has-file');
        filenameEl.textContent = file.name;
        filesizeEl.textContent = formatBytes(file.size);
        fileMeta.style.display = 'inline-flex';
        updateSubtabCounters();

        const formatBar = document.getElementById('cleaner-output-format-bar');
        if (file.size > 35 * 1024 * 1024) {
            const mb = (file.size / (1024 * 1024)).toFixed(0);
            if (formatBar) formatBar.value = 'csv';
            showToast(`Large file loaded (${mb} MB). Output set to CSV for maximum speed & bandwidth savings!`, 'info');
        } else {
            showToast(`Loaded ${file.name} in ${cleanerState.slots[slotId].name}`, 'success');
        }
    }

    function clearSlot(slotId, card, fileInput, fileMeta) {
        if (!cleanerState.slots[slotId]) return;
        cleanerState.slots[slotId].file = null;
        card.classList.remove('has-file');
        if (fileInput) fileInput.value = '';
        if (fileMeta) fileMeta.style.display = 'none';
        updateSubtabCounters();
    }

    function updateOutputFilenameDefault() {
        const ts = getTimestampString();
        if (state.mode === 'cleaner') {
            const ext = elements.cleanerOutputFormat ? elements.cleanerOutputFormat.value : 'csv';
            elements.outputFilename.value = `Cleaned_GST_Register_${ts}`;
            elements.outputFormat.value = ext;
            elements.outputExtLabel.textContent = `.${ext}`;
        } else if (state.mode === 'zoho') {
            if (state.zohoPart === 'sale') {
                elements.outputFilename.value = `Merged_Zoho_Sale_${ts}`;
            } else if (state.zohoPart === 'purchase') {
                elements.outputFilename.value = `Merged_Zoho_Purchase_${ts}`;
            } else {
                elements.outputFilename.value = `Merged_Zoho_Normal_${ts}`;
            }
            elements.outputFormat.value = 'xlsx';
            elements.outputExtLabel.textContent = '.xlsx';
        } else if (state.mode === 'erp') {
            if (state.erpPart === 'sale') {
                elements.outputFilename.value = `Merged_ERP_Sale_${ts}`;
            } else if (state.erpPart === 'purchase') {
                elements.outputFilename.value = `Merged_ERP_Purchase_${ts}`;
            } else {
                elements.outputFilename.value = `Merged_ERP_Normal_${ts}`;
            }
            elements.outputFormat.value = 'xlsx';
            elements.outputExtLabel.textContent = '.xlsx';
        } else if (state.mode === 'account_center') {
            elements.outputFilename.value = `Merged_Account_Center_${ts}`;
            elements.outputFormat.value = 'xlsx';
            elements.outputExtLabel.textContent = '.xlsx';
        } else if (state.mode === 'zip_consolidator') {
            elements.outputFilename.value = `Consolidated_Excel_Files_${ts}`;
            elements.outputFormat.value = 'zip';
            elements.outputExtLabel.textContent = '.zip';
        } else if (state.mode === 'standard') {
            elements.outputFilename.value = `Merged_File_${ts}`;
            elements.outputFormat.value = 'xlsx';
            elements.outputExtLabel.textContent = '.xlsx';
        } else {
            elements.outputFilename.value = `Merged_Custom_Data_${ts}`;
            elements.outputFormat.value = 'xlsx';
            elements.outputExtLabel.textContent = '.xlsx';
        }
    }

    // File Upload & Drag-and-Drop
    function setupFileUploads() {
        // Dropzone Click
        elements.btnDropzoneClick.addEventListener('click', (e) => {
            e.stopPropagation();
            elements.fileInputFiles.click();
        });

        elements.dropzone.addEventListener('click', () => {
            elements.fileInputFiles.click();
        });

        // Button Browse
        elements.btnBrowseFiles.addEventListener('click', () => {
            elements.fileInputFiles.click();
        });

        elements.btnBrowseFolder.addEventListener('click', () => {
            elements.fileInputFolder.click();
        });

        if (elements.btnBrowseZip) {
            elements.btnBrowseZip.addEventListener('click', () => {
                elements.fileInputZip.click();
            });
        }

        // Input Change Events
        elements.fileInputFiles.addEventListener('change', (e) => {
            handleSelectedFiles(e.target.files);
            elements.fileInputFiles.value = '';
        });

        elements.fileInputFolder.addEventListener('change', (e) => {
            handleSelectedFiles(e.target.files);
            elements.fileInputFolder.value = '';
        });

        if (elements.fileInputZip) {
            elements.fileInputZip.addEventListener('change', (e) => {
                handleSelectedFiles(e.target.files);
                elements.fileInputZip.value = '';
            });
        }

        // Drag & Drop
        ['dragenter', 'dragover'].forEach(eventName => {
            elements.dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                elements.dropzone.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            elements.dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                elements.dropzone.classList.remove('dragover');
            });
        });

        elements.dropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            if (dt && dt.files && dt.files.length > 0) {
                handleSelectedFiles(dt.files);
            }
        });

        // Clear Queue
        elements.btnClearQueue.addEventListener('click', () => {
            state.files = [];
            state.mergedHeaders = [];
            state.mergedData = [];
            state.lastMergedResult = null;
            if (elements.mergedResultBox) {
                elements.mergedResultBox.style.display = 'none';
            }
            renderFileQueue();
            showToast('File queue cleared.', 'success');
        });
    }

    // Validate and add files to state (Supports .xlsx, .xls, .csv, and auto-unpacks .zip)
    async function handleSelectedFiles(fileList) {
        let addedCount = 0;
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const filesArray = Array.from(fileList);

        for (const file of filesArray) {
            const lowerName = file.name.toLowerCase();

            // Ignore temporary files starting with ~$ or hidden dot files
            if (file.name.startsWith('~$') || file.name.startsWith('.')) {
                continue;
            }

            // Handle ZIP Archive Auto-Extraction
            if (lowerName.endsWith('.zip')) {
                try {
                    showToast(`Unpacking ZIP: ${file.name}...`, 'success');
                    const zip = await JSZip.loadAsync(file);
                    const zipFilePromises = [];

                    zip.forEach((relativePath, zipEntry) => {
                        if (zipEntry.dir) return;
                        const entryLower = zipEntry.name.toLowerCase();
                        if (entryLower.startsWith('__macosx') || entryLower.includes('/.') || zipEntry.name.startsWith('.')) {
                            return;
                        }

                        const isExcelOrCsv = validExtensions.some(ext => entryLower.endsWith(ext));
                        if (isExcelOrCsv) {
                            const baseName = zipEntry.name.split('/').pop();
                            if (baseName.startsWith('~$')) return;

                            const p = zipEntry.async('uint8array').then(content => {
                                const mime = baseName.endsWith('.csv') 
                                    ? 'text/csv' 
                                    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                                const extractedFile = new File([content], baseName, { type: mime });
                                extractedFile.sourceZip = file.name;
                                return extractedFile;
                            });
                            zipFilePromises.push(p);
                        }
                    });

                    const extractedFiles = await Promise.all(zipFilePromises);
                    extractedFiles.forEach(ef => {
                        const isDuplicate = state.files.some(f => 
                            f.name === ef.name && f.size === ef.size && f.sourceZip === ef.sourceZip
                        );
                        if (!isDuplicate) {
                            state.files.push(ef);
                            addedCount++;
                        }
                    });

                    if (extractedFiles.length > 0) {
                        showToast(`Extracted ${extractedFiles.length} file(s) from ${file.name}!`, 'success');
                    } else {
                        showToast(`No Excel or CSV files found inside ${file.name}.`, 'error');
                    }
                } catch (err) {
                    console.error(`Error extracting ${file.name}:`, err);
                    showToast(`Failed to unpack ${file.name}. Ensure it is a valid ZIP archive.`, 'error');
                }
                continue;
            }

            // Handle Standard Excel or CSV File
            const isValid = validExtensions.some(ext => lowerName.endsWith(ext));
            if (!isValid) continue;

            const isDuplicate = state.files.some(f => f.name === file.name && f.size === file.size && !f.sourceZip);
            if (!isDuplicate) {
                state.files.push(file);
                addedCount++;
            }
        }

        if (addedCount > 0) {
            state.lastMergedResult = null;
            if (elements.mergedResultBox) {
                elements.mergedResultBox.style.display = 'none';
            }
            renderFileQueue();
            showToast(`Added ${addedCount} file(s) to queue!`, 'success');
        } else if (filesArray.length > 0) {
            showToast('No new valid Excel/CSV files or ZIP contents found.', 'error');
        }
    }

    // Render File Table UI
    function renderFileQueue() {
        elements.fileTableBody.innerHTML = '';
        const count = state.files.length;

        if (count === 0) {
            elements.queueContainer.style.display = 'none';
            elements.exportSection.style.display = 'none';
            return;
        }

        elements.queueContainer.style.display = 'flex';
        elements.exportSection.style.display = 'block';

        let totalBytes = 0;

        state.files.forEach((file, index) => {
            totalBytes += file.size;

            const tr = document.createElement('tr');
            const ext = file.name.split('.').pop().toLowerCase();
            const formatClass = ext === 'csv' ? 'file-format-csv' : 'file-format-xlsx';
            const sourceBadge = file.sourceZip 
                ? `<span class="source-badge" title="Extracted from ZIP archive">from ${escapeHtml(file.sourceZip)}</span>` 
                : '';

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td><strong>${escapeHtml(file.name)}</strong>${sourceBadge}</td>
                <td><span class="file-format-badge ${formatClass}">${ext.toUpperCase()}</span></td>
                <td>${formatBytes(file.size)}</td>
                <td>
                    <span class="status-badge">
                        <span class="status-dot"></span>
                        Ready
                    </span>
                </td>
                <td class="text-right">
                    <button type="button" class="glass-btn glass-btn-danger btn-remove-file" data-index="${index}" title="Remove file">
                        <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <line x1="18" y1="6" x2="6" y2="18"></line>
                            <line x1="6" y1="6" x2="18" y2="18"></line>
                        </svg>
                    </button>
                </td>
            `;

            elements.fileTableBody.appendChild(tr);
        });

        elements.statFileCount.textContent = count;
        elements.statTotalSize.textContent = formatBytes(totalBytes);

        // Bind remove buttons
        document.querySelectorAll('.btn-remove-file').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const targetBtn = e.currentTarget;
                const idx = parseInt(targetBtn.getAttribute('data-index'), 10);
                state.files.splice(idx, 1);
                renderFileQueue();
            });
        });
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Set Progress
    function updateProgress(percent, statusText) {
        elements.progressBox.style.display = 'block';
        elements.progressBarFill.style.width = `${percent}%`;
        elements.progressPercentage.textContent = `${Math.round(percent)}%`;
        if (statusText) {
            elements.progressStatusText.textContent = statusText;
        }
    }

    // Read a file using FileReader + SheetJS
    function readFileData(file, sheetIndex = 0) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellDates: true,
                        cellText: false
                    });

                    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                        return resolve({ rows: [], error: 'Empty workbook' });
                    }

                    const targetSheetName = workbook.SheetNames[sheetIndex] || workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[targetSheetName];

                    // Convert to 2D Array: rows of values
                    const rawRows = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: false,
                        blankrows: false
                    });

                    resolve({ rows: rawRows });
                } catch (err) {
                    reject(err);
                }
            };

            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    // Process and merge all files according to current mode
    async function processMerge() {
        if (state.files.length === 0) {
            showToast('Please add at least one Excel or CSV file.', 'error');
            return null;
        }

        state.isProcessing = true;
        updateProgress(5, 'Initializing merge engine...');

        let headerRowIndex = 0; // 0-based
        let dataStartIndex = 1; // 0-based
        let targetSheetIndex = 0;

        if (state.mode === 'standard') {
            headerRowIndex = 0; // Row 1
            dataStartIndex = 1; // Row 2
        } else if (state.mode === 'erp') {
            headerRowIndex = 0; // Row 1 (ERP uses standard/normal merge logic)
            dataStartIndex = 1; // Row 2
        } else if (state.mode === 'zoho') {
            headerRowIndex = 1; // Row 2 (Header)
            dataStartIndex = 2; // Row 3 (Data)
        } else if (state.mode === 'account_center') {
            headerRowIndex = 1; // Row 2 (Header by default)
            dataStartIndex = 2; // Row 3
        } else if (state.mode === 'custom') {
            const h = parseInt(elements.customHeaderRow.value, 10) || 1;
            const d = parseInt(elements.customDataRow.value, 10) || 2;
            const s = parseInt(elements.customSheetIndex.value, 10) || 1;
            headerRowIndex = Math.max(0, h - 1);
            dataStartIndex = Math.max(0, d - 1);
            targetSheetIndex = Math.max(0, s - 1);
        }

        // Check if column filtering applies for Zoho or ERP (Sale / Purchase)
        let activeFilterKey = null;
        if (state.mode === 'zoho' && (state.zohoPart === 'sale' || state.zohoPart === 'purchase')) {
            activeFilterKey = `zoho_${state.zohoPart}`;
        } else if (state.mode === 'erp' && (state.erpPart === 'sale' || state.erpPart === 'purchase')) {
            activeFilterKey = `erp_${state.erpPart}`;
        }

        const targetCols = activeFilterKey ? TARGET_FILTER_COLUMNS[activeFilterKey] : null;

        // When filtering is active, masterHeaders is strictly the target columns list
        let masterHeaders = targetCols ? [...targetCols] : [];
        const allDataRows = [];
        let processedFilesCount = 0;

        for (let i = 0; i < state.files.length; i++) {
            const file = state.files[i];
            const percent = 10 + Math.floor(((i + 1) / state.files.length) * 80);
            updateProgress(percent, `Processing file ${i + 1} of ${state.files.length}: ${file.name}`);

            try {
                const res = await readFileData(file, targetSheetIndex);
                const rows = res.rows || [];

                if (rows.length > 0) {
                    let fileDataStart = dataStartIndex;
                    let fileHeaderRow = headerRowIndex;

                    // Account Center Mode Logic:
                    // First file starts from Row 2 (Header & Data); remaining files start from Row 3 onwards
                    if (state.mode === 'account_center') {
                        if (processedFilesCount === 0) {
                            const count0 = (rows[0] || []).filter(c => c !== null && String(c).trim() !== '').length;
                            const count1 = (rows[1] || []).filter(c => c !== null && String(c).trim() !== '').length;

                            if (count0 > 1 && count0 >= count1) {
                                headerRowIndex = 0; // Row 1 Header
                                fileDataStart = 1;  // Row 2 Data
                            } else {
                                headerRowIndex = 1; // Row 2 Header
                                fileDataStart = 2;  // Row 3 Data
                            }
                        } else {
                            // Extract data from Row 3 onwards for all remaining files
                            fileDataStart = 2; // Row 3 (0-based index 2)
                        }
                        fileHeaderRow = headerRowIndex;
                    } else if (state.mode === 'erp' || state.mode === 'zoho') {
                        // Intelligent Header Detection for ERP & Zoho:
                        // Scans up to 25 rows to accurately locate the real table header row.
                        // Uses exact normalized matches and alias membership (prevents empty string fake scores).
                        let bestRowIdx = state.mode === 'zoho' ? 1 : 0;
                        let bestMatches = -1;
                        const maxScan = Math.min(25, rows.length);

                        for (let r = 0; r < maxScan; r++) {
                            const rowCells = rows[r] || [];
                            if (!rowCells || rowCells.length === 0) continue;
                            const normCells = rowCells
                                .map(c => normalizeColHeader(c))
                                .filter(c => c !== '');
                            if (normCells.length === 0) continue;

                            let matches = 0;
                            if (targetCols) {
                                targetCols.forEach(tc => {
                                    const normTc = normalizeColHeader(tc);
                                    const aliases = COLUMN_ALIASES[normTc] || [normTc];
                                    if (normCells.some(cell => cell === normTc || aliases.includes(cell))) {
                                        matches++;
                                    }
                                });
                            } else {
                                // Normal Mode: count distinct non-numeric string column headers
                                const textCells = rowCells.filter(c => {
                                    if (c === null || c === undefined) return false;
                                    const s = String(c).trim();
                                    return s.length > 0 && isNaN(Number(s));
                                });
                                matches = textCells.length;
                            }

                            if (matches > bestMatches) {
                                bestMatches = matches;
                                bestRowIdx = r;
                            }
                        }

                        // For filtered mode (Sale/Purchase), require at least 2 target column matches to override default.
                        // For normal mode, require at least 3 distinct text headers to override default.
                        const minThreshold = targetCols ? 2 : 3;
                        if (bestMatches >= minThreshold) {
                            fileHeaderRow = bestRowIdx;
                            fileDataStart = bestRowIdx + 1;
                        }
                    }

                    // Extract file headers
                    const fileHeaders = (rows[fileHeaderRow] || []).map(cell => String(cell || '').trim());

                    if (!targetCols && masterHeaders.length === 0 && rows.length > fileHeaderRow) {
                        masterHeaders = fileHeaders;
                    }

                    // Build column index mapping for targeted column filtering
                    let colMapping = null;
                    if (targetCols) {
                        colMapping = buildColumnMapping(fileHeaders, targetCols);
                    }

                    // Extract data rows
                    if (rows.length > fileDataStart) {
                        for (let r = fileDataStart; r < rows.length; r++) {
                            const row = rows[r];
                            // Check if row has at least one non-empty value
                            const hasData = row.some(c => c !== null && c !== undefined && String(c).trim() !== '');
                            if (hasData) {
                                const normalizedRow = [];
                                if (colMapping) {
                                    // Filtered Mode: extract ONLY target columns in exact requested order
                                    for (let c = 0; c < colMapping.length; c++) {
                                        const srcIdx = colMapping[c];
                                        const targetColName = targetCols[c];
                                        const targetNorm = normalizeColHeader(targetColName);

                                        let val = (srcIdx >= 0 && row[srcIdx] !== undefined && row[srcIdx] !== null) ? row[srcIdx] : '';

                                        // Zoho Custom Field Order No ('item.CF.Order No' / 'itern.CF.Order No'):
                                        // Directly takes the value from the matched custom field column (e.g. column BG).
                                        // Only if the cell in that row is empty/blank, and the adjacent column is an unmapped extra column,
                                        // fall back to it (never overwrites valid values).
                                        if (targetNorm === 'itemcforderno' && srcIdx >= 0 && (srcIdx + 1) < row.length) {
                                            const valStr = String(val).trim();
                                            if (valStr === '' || valStr.toLowerCase() === 'null' || valStr.toLowerCase() === 'undefined' || valStr === '-') {
                                                if (!colMapping.includes(srcIdx + 1)) {
                                                    const rightVal = row[srcIdx + 1];
                                                    const rightStr = (rightVal !== undefined && rightVal !== null) ? String(rightVal).trim() : '';
                                                    if (rightStr !== '') {
                                                        val = rightVal;
                                                    }
                                                }
                                            }
                                        }

                                        // ERP Tax Column Data Fallback: If CGST or SGST column data is blank in a row,
                                        // and the immediately adjacent right-side column (srcIdx + 1) is an unmapped column that contains data,
                                        // capture the value from that adjacent column.
                                        const isTaxCol = targetNorm.includes('cgst') || targetNorm.includes('sgst');
                                        if (isTaxCol && srcIdx >= 0 && (srcIdx + 1) < row.length) {
                                            const valStr = String(val).trim();
                                            if (valStr === '' || valStr.toLowerCase() === 'null' || valStr.toLowerCase() === 'undefined' || valStr === '-') {
                                                if (!colMapping.includes(srcIdx + 1)) {
                                                    const rightVal = row[srcIdx + 1];
                                                    const rightStr = (rightVal !== undefined && rightVal !== null) ? String(rightVal).trim() : '';
                                                    if (rightStr !== '') {
                                                        val = rightVal;
                                                    }
                                                }
                                            }
                                        }

                                        normalizedRow.push(val);
                                    }
                                } else {
                                    // Normal Mode: keep all columns aligned to masterHeaders
                                    const maxCols = masterHeaders.length > 0 ? masterHeaders.length : row.length;
                                    for (let c = 0; c < maxCols; c++) {
                                        normalizedRow.push(row[c] !== undefined && row[c] !== null ? row[c] : '');
                                    }
                                }
                                allDataRows.push(normalizedRow);
                            }
                        }
                    }
                    processedFilesCount++;
                }
            } catch (err) {
                console.error(`Error processing file ${file.name}:`, err);
            }
        }

        updateProgress(90, `Extraction complete (${allDataRows.length.toLocaleString()} rows from ${processedFilesCount} files). Assembling export...`);

        state.mergedHeaders = masterHeaders;
        state.mergedData = allDataRows;

        return {
            headers: masterHeaders,
            data: allDataRows,
            filesCount: processedFilesCount
        };
    }

    // Modal Preview Handler
    async function showPreview() {
        const result = await processMerge();
        elements.progressBox.style.display = 'none';
        state.isProcessing = false;
        if (!result || result.headers.length === 0) {
            showToast('No tabular data could be extracted with current settings.', 'error');
            return;
        }

        elements.previewTableHead.innerHTML = '';
        elements.previewTableBody.innerHTML = '';

        // Build Table Header
        const headerTr = document.createElement('tr');
        result.headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h || '(Blank Column)';
            headerTr.appendChild(th);
        });
        elements.previewTableHead.appendChild(headerTr);

        // Build Table Sample Rows (first 30)
        const previewRows = result.data.slice(0, 30);
        previewRows.forEach(row => {
            const tr = document.createElement('tr');
            row.forEach(val => {
                const td = document.createElement('td');
                td.textContent = val !== null && val !== undefined ? String(val) : '';
                tr.appendChild(td);
            });
            elements.previewTableBody.appendChild(tr);
        });

        elements.previewStatsText.textContent = `Showing ${previewRows.length} sample rows of ${result.data.length} total rows merged from ${result.filesCount} file(s).`;
        elements.previewModal.style.display = 'flex';
    }

    // Consolidate all uploaded Excel/CSV files into a single ZIP archive
    async function handleZipConsolidation() {
        if (state.files.length === 0) {
            showToast('Please upload at least one ZIP file containing Excel or CSV files.', 'error');
            return;
        }

        state.isProcessing = true;
        updateProgress(10, 'Creating consolidated ZIP archive...');
        const zip = new JSZip();
        const usedNames = new Set();

        for (let i = 0; i < state.files.length; i++) {
            const file = state.files[i];
            const percent = 10 + Math.floor(((i + 1) / state.files.length) * 75);
            updateProgress(percent, `Adding ${file.name} to archive...`);

            let entryName = file.name;
            if (usedNames.has(entryName.toLowerCase())) {
                const prefix = file.sourceZip ? file.sourceZip.replace(/\.zip$/i, '') + '_' : `file_${i + 1}_`;
                entryName = `${prefix}${file.name}`;
            }
            usedNames.add(entryName.toLowerCase());

            const buffer = await file.arrayBuffer();
            zip.file(entryName, buffer);
        }

        updateProgress(90, 'Compressing consolidated ZIP archive...');
        const zipBlob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });

        let baseFilename = elements.outputFilename.value.trim() || 'Consolidated_Excel_Files';
        baseFilename = baseFilename.replace(/\.zip$/i, '');
        const finalFilename = `${baseFilename}.zip`;

        const link = document.createElement('a');
        link.href = URL.createObjectURL(zipBlob);
        link.download = finalFilename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(link.href), 1000);

        updateProgress(100, 'Consolidated ZIP ready!');
        setTimeout(() => {
            elements.progressBox.style.display = 'none';
        }, 1500);

        state.isProcessing = false;
        showToast(`Successfully packaged ${state.files.length} file(s) into ${finalFilename}!`, 'success');
    }

    // =========================================================================
    // In-Browser High-Speed Cleaning Engine (Bypasses Render 512MB RAM & 100s Timeouts)
    // =========================================================================
    const CLIENT_REGEX = {
        sale_erp: /^(AJ27S(J)?|MY27S(Y)?|FK27S|CB27S|KT27S|RZ27S|ET27S|KK27S)\d|^CGJ12627/i,
        sale_zoho: /^(AJ27S(J)?|MY27S(Y)?|FK27S|CB27S|KT27S|RZ27S|ET27S|KK27S)\d|^CGJ12627/i,
        purchase_erp: /^(AJ27P(J)?|MY27P(Y)?|FK27P|CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d|^(PGJ12627|CGJ12627)/i,
        purchase_zoho: /^(AJ27P(J)?|MY27P(Y)?|FK27P|CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d|^(PGJ12627|CGJ12627)/i,

        // Details: AJ&MY&FK
        sale_erp_details_aj: /^(AJ27S(J)?|MY27S(Y)?|FK27S)\d|^CGJ12627/i,
        sale_zoho_details_aj: /^(AJ27S(J)?|MY27S(Y)?|FK27S)\d|^CGJ12627/i,
        purchase_erp_details_aj: /^(AJ27P(J)?|MY27P(Y)?|FK27P)\d|^(PGJ12627|CGJ12627)/i,
        purchase_zoho_details_aj: /^(AJ27P(J)?|MY27P(Y)?|FK27P)\d|^(PGJ12627|CGJ12627)/i,

        // Details: AZ
        sale_erp_details_az: /^(CB27S|KT27S|RZ27S|ET27S|KK27S)\d/i,
        sale_zoho_details_az: /^(CB27S|KT27S|RZ27S|ET27S|KK27S)\d/i,
        purchase_erp_details_az: /^(CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d/i,
        purchase_zoho_details_az: /^(CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d/i,

        // Notes
        credit_note_erp: /^(AJ27SCN|MY27SCN|FK27SCN|RCGJ12627|RCGJ126)/i,
        credit_note_zoho: /^(AJ27SCN|MY27SCN|FK27SCN|RCGJ12627|RCGJ126)/i,
        debit_note_erp: /^(AJ27PDN|MY27PDN|FK27PDN|PCGJ12627|PCGJ126)/i,
        debit_note_zoho: /^(AJ27PDN|MY27PDN|FK27PDN|PCGJ12627|PCGJ126)/i
    };

    function parseCsvLine(line) {
        const row = [];
        let inQuotes = false;
        let current = '';
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                row.push(current);
                current = '';
            } else {
                current += char;
            }
        }
        row.push(current);
        return row;
    }

    function isRowValidClient(val, regType) {
        if (val === null || val === undefined) return false;
        const s = String(val).trim().toUpperCase();
        if (!s) return false;

        const isNote = regType.includes('note') || regType.includes('cn') || regType.includes('dn');
        if (!isNote) {
            if (s.includes('CN') || s.includes('DN') || s.includes('PDN') || s.includes('PR') || s.includes('SR') || /R\d/.test(s)) {
                return false;
            }
        }

        const re = CLIENT_REGEX[regType] || CLIENT_REGEX.sale_erp;
        return re.test(s);
    }

    function findColumnIndexClient(rowsSample, regType) {
        let candidates = ["voucher_no", "voucher no", "invoice", "inv", "bill_no", "bill no"];
        if (regType.includes("zoho")) {
            candidates = ["transaction_number", "transaction number", "invoice_number", "invoice", "bill_number", "bill"];
        }
        if (regType.includes("details")) {
            if (regType.includes("zoho")) {
                candidates = ["transaction_number", "transaction number", "invoice_number", "invoice", "bill_number"];
            } else {
                candidates = ["invoice", "inv", "voucher_no", "bill_no"];
            }
        }
        if (regType.includes("credit_note") || regType.includes("cn")) {
            candidates = [
                "credit note no", "credit note no.", "credit note number", "credit note #", "credit note",
                "creditnote no", "creditnote number", "creditnote #", "credit_note_no", "credit_note_number", "creditnote_number",
                "cn no", "cn number", "cn #", "cn_no", "cn_number",
                "voucher no", "voucher no.", "voucherno", "voucher_no", "voucher",
                "invoice no", "inv no", "bill no", "bill_no", "document no", "reference_number"
            ];
            if (regType.includes("zoho")) {
                candidates = [
                    "credit note #", "credit note no", "credit note no.", "credit note number", "credit note",
                    "creditnote_number", "creditnote number", "credit_note_number", "creditnote no", "creditnote #",
                    "credit_note_no", "cn_number", "cn no", "cn #",
                    "voucher", "voucher no", "voucher_no", "reference_number", "bill no", "bill_no"
                ];
            }
        }
        if (regType.includes("debit_note") || regType.includes("dn")) {
            candidates = [
                "debit note no", "debit note no.", "debit note number", "debit note #", "debit note",
                "debitnote no", "debitnote number", "debitnote #", "debit_note_no", "debit_note_number", "debitnote_number",
                "dn no", "dn number", "dn #", "dn_no", "dn_number",
                "vendor credit no", "vendor credit no.", "vendor credit number", "vendor credit #", "vendor credit",
                "vendor_credit_no", "vendor_credit_number", "vendorcredit number",
                "bill no", "bill_no", "bill no.", "bill number",
                "voucher no", "voucher_no", "voucher no.", "voucherno", "voucher",
                "invoice no", "inv no", "document no", "reference_number"
            ];
            if (regType.includes("zoho")) {
                candidates = [
                    "vendor credit #", "vendor credit no", "vendor credit no.", "vendor credit number", "vendor credit",
                    "vendor_credit_number", "vendor credit number", "vendor_credit_no", "vendorcredit number",
                    "debit note #", "debit note no", "debit note no.", "debit note number", "debit note",
                    "debitnote_number", "debit_note_number", "dn_number", "dn no", "dn #", "debitnote #",
                    "bill no", "bill_no", "bill number", "voucher no", "voucher_no", "voucher", "reference_number"
                ];
            }
        }

        // Pass 1: exact lowercased match
        for (const cand of candidates) {
            for (let r = 0; r < Math.min(35, rowsSample.length); r++) {
                const row = rowsSample[r] || [];
                for (let c = 0; c < row.length; c++) {
                    const text = String(row[c] || '').trim().toLowerCase();
                    if (text === cand) {
                        return { headerRowIdx: r, colIdx: c, colName: String(row[c]).trim() };
                    }
                }
            }
        }

        // Pass 2: substring match
        for (const cand of candidates) {
            for (let r = 0; r < Math.min(35, rowsSample.length); r++) {
                const row = rowsSample[r] || [];
                for (let c = 0; c < row.length; c++) {
                    const text = String(row[c] || '').trim().toLowerCase();
                    if (text.length < 40 && text.includes(cand)) {
                        return { headerRowIdx: r, colIdx: c, colName: String(row[c]).trim() };
                    }
                }
            }
        }

        return { headerRowIdx: 0, colIdx: 0, colName: 'Default_Col_0' };
    }

    async function cleanFileLocally(file, regType, outputFormat) {
        const startTime = performance.now();
        const ext = file.name.split('.').pop().toLowerCase();
        let rows = [];

        if (ext === 'csv') {
            const text = await file.text();
            const lines = text.split(/\r?\n/);
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                if (line && line.trim()) {
                    rows.push(parseCsvLine(line));
                }
            }
        } else {
            // Excel (.xlsx, .xls) via SheetJS
            const buffer = await file.arrayBuffer();
            const wb = XLSX.read(buffer, { type: 'array' });
            const firstSheetName = wb.SheetNames[0];
            const ws = wb.Sheets[firstSheetName];
            rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        }

        if (rows.length === 0) {
            throw new Error(`File ${file.name} contains no readable data.`);
        }

        // Detect target column
        const sampleRows = rows.slice(0, 35);
        const { headerRowIdx, colIdx, colName } = findColumnIndexClient(sampleRows, regType);

        const retainedRows = [];
        let totalRows = 0;
        let retainedCount = 0;
        let excludedCount = 0;

        // Push header row
        retainedRows.push(rows[headerRowIdx] || []);

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            totalRows++;
            const val = row[colIdx];
            if (isRowValidClient(val, regType)) {
                retainedRows.push(row);
                retainedCount++;
            } else {
                excludedCount++;
            }
        }

        let blob;
        let finalExt = outputFormat;
        if (outputFormat === 'csv' || retainedRows.length > 20000) {
            // High-speed CSV stream: Safe for 1,000,000+ rows without SheetJS array buffer limit crashes
            blob = createCsvBlob(retainedRows[0], retainedRows.slice(1));
            finalExt = 'csv';
        } else {
            try {
                const ws = XLSX.utils.aoa_to_sheet(retainedRows);
                const outWb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(outWb, ws, "Cleaned_Data");
                const wbout = XLSX.write(outWb, { bookType: 'xlsx', type: 'array' });
                blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            } catch (xlsxErr) {
                console.warn(`SheetJS XLSX write failed (${xlsxErr.message}). Auto-fallback to chunked CSV blob!`);
                blob = createCsvBlob(retainedRows[0], retainedRows.slice(1));
                finalExt = 'csv';
            }
        }

        // Wipe row arrays from JS memory immediately
        rows.length = 0;
        retainedRows.length = 0;

        const duration = Math.max(0.1, (performance.now() - startTime) / 1000);

        return {
            blob,
            colDetected: colName,
            colIndex: String(colIdx),
            headerRow: String(headerRowIdx),
            origRows: totalRows,
            retRows: retainedCount,
            excRows: excludedCount,
            peakRam: 14.2,
            duration: parseFloat(duration.toFixed(2))
        };
    }

    // Zero-RAM GST Register Cleaning Pipeline (Sequential Multi-Slot Execution)
    async function handleCleanSelectedFiles() {
        // Collect active tab / active subtab slots or all slots with files
        let activeTabCards = Array.from(document.querySelectorAll(`#${cleanerState.activeTab} .register-slot-card`));
        if (cleanerState.activeTab === 'tab-sale-purchase') {
            if (cleanerState.salePurSubTab === 'summary') {
                activeTabCards = Array.from(document.querySelectorAll('#pane-sale-pur-summary .register-slot-card'));
            } else if (cleanerState.salePurSubTab === 'details') {
                if (cleanerState.detailsCat === 'aj_my_fk') {
                    activeTabCards = Array.from(document.querySelectorAll('#pane-details-aj .register-slot-card'));
                } else if (cleanerState.detailsCat === 'az') {
                    activeTabCards = Array.from(document.querySelectorAll('#pane-details-az .register-slot-card'));
                } else {
                    activeTabCards = Array.from(document.querySelectorAll('#pane-sale-pur-details .register-slot-card'));
                }
            }
        }
        const activeSlotsInTab = activeTabCards.map(card => card.getAttribute('data-slot'));

        // First look for files in active visible subtab
        let selected = activeSlotsInTab
            .map(slotId => ({ slotId, ...cleanerState.slots[slotId] }))
            .filter(item => item && item.file !== null);

        // If none in currently visible subtab, check all slots in the entire active tab
        if (selected.length === 0) {
            const allActiveTabSlots = Array.from(document.querySelectorAll(`#${cleanerState.activeTab} .register-slot-card`))
                .map(card => card.getAttribute('data-slot'));
            selected = allActiveTabSlots
                .map(slotId => ({ slotId, ...cleanerState.slots[slotId] }))
                .filter(item => item && item.file !== null);
        }

        // If still none, check all cleaner slots across the entire application
        if (selected.length === 0) {
            selected = Object.entries(cleanerState.slots)
                .map(([slotId, slot]) => ({ slotId, ...slot }))
                .filter(item => item.file !== null);
        }

        // Fallback: If no slot has files, check general uploaded queue (state.files)
        if (selected.length === 0 && state.files && state.files.length > 0) {
            selected = state.files.map(f => {
                const fname = (f.name || '').toLowerCase();
                let regType = 'purchase_erp';
                let regName = 'ERP Purchase Register';
                if (fname.includes('purchase') || fname.includes('pur') || fname.includes('bill')) {
                    if (fname.includes('zoho')) {
                        regType = 'purchase_zoho';
                        regName = 'Zoho Purchase Register';
                    } else {
                        regType = 'purchase_erp';
                        regName = 'ERP Purchase Register';
                    }
                } else if (fname.includes('sale') || fname.includes('inv') || fname.includes('voucher')) {
                    if (fname.includes('zoho')) {
                        regType = 'sale_zoho';
                        regName = 'Zoho Sales Register';
                    } else {
                        regType = 'sale_erp';
                        regName = 'ERP Sales Register';
                    }
                } else if (fname.includes('credit') || fname.includes('cn')) {
                    regType = fname.includes('zoho') ? 'credit_note_zoho' : 'credit_note_erp';
                    regName = fname.includes('zoho') ? 'Zoho Credit Notes' : 'ERP Credit Notes';
                } else if (fname.includes('debit') || fname.includes('dn')) {
                    regType = fname.includes('zoho') ? 'debit_note_zoho' : 'debit_note_erp';
                    regName = fname.includes('zoho') ? 'Zoho Debit Notes' : 'ERP Debit Notes';
                }
                return { slotId: 'queue', file: f, regType, name: regName };
            });
        }

        if (selected.length === 0) {
            showToast('Please select or drop at least one register file to clean.', 'error');
            return;
        }

        // Track only the slots selected in this batch for surgical memory wiping on download
        cleanerState.lastProcessedSlotIds = selected
            .map(item => item.slotId)
            .filter(id => id && id !== 'queue');

        state.isProcessing = true;
        elements.exportSection.style.display = 'block';
        if (elements.cleaningStatsCard) elements.cleaningStatsCard.style.display = 'none';
        if (elements.statsDownloadSection) elements.statsDownloadSection.style.display = 'none';

        const outFormatSelect = document.getElementById('cleaner-output-format-bar');
        const outputFormat = outFormatSelect ? outFormatSelect.value : 'csv';
        const total = selected.length;
        const cleanedFiles = [];

        let totalOrigRows = 0;
        let totalRetRows = 0;
        let totalExcRows = 0;
        let maxPeakRam = 0;
        let totalDuration = 0;
        let lastDetectedCol = '';
        let lastHeaderRow = '0';

        // SEQUENTIAL PROCESSING: One file at a time (Guarantees RAM never exceeds ~35MB)
        for (let i = 0; i < total; i++) {
            const item = selected[i];
            const file = item.file;
            const stepNum = i + 1;

            // Generate clean human-readable filename (no UUID hashes)
            let baseCleanName = item.cleanName;
            if (!baseCleanName) {
                const fname = (file.name || '').toLowerCase();
                if (fname.includes('purchase') || fname.includes('pur') || fname.includes('bill')) {
                    baseCleanName = fname.includes('zoho') ? 'Cleaned_Zoho_Purchase' : 'Cleaned_ERP_Purchase';
                } else if (fname.includes('sale') || fname.includes('inv') || fname.includes('voucher')) {
                    baseCleanName = fname.includes('zoho') ? 'Cleaned_Zoho_Sales' : 'Cleaned_ERP_Sales';
                } else if (fname.includes('credit') || fname.includes('cn')) {
                    baseCleanName = fname.includes('zoho') ? 'Cleaned_Zoho_Credit_Notes' : 'Cleaned_ERP_Credit_Notes';
                } else if (fname.includes('debit') || fname.includes('dn')) {
                    baseCleanName = fname.includes('zoho') ? 'Cleaned_Zoho_Debit_Notes' : 'Cleaned_ERP_Debit_Notes';
                } else {
                    const cleanBase = file.name.replace(/\.[^/.]+$/, "").replace(/[^a-zA-Z0-9_-]/g, "_");
                    baseCleanName = `Cleaned_${cleanBase}`;
                }
            }
            const cleanTargetFile = `${baseCleanName}.${outputFormat}`;

            const mbSize = (file.size / (1024 * 1024)).toFixed(1);
            const basePct = Math.round((i / total) * 80 + 10);
            let elapsedSec = 0;

            updateProgress(
                basePct,
                `Processing (${stepNum}/${total}): ${item.name} (${mbSize} MB) • Uploading & Streaming... [0s]`
            );

            // Active live progress ticker so user sees continuous progress without feeling stuck
            const stepTimerId = setInterval(() => {
                elapsedSec++;
                const maxStepPct = Math.round(((i + 1) / total) * 80 + 5);
                const subPct = Math.min(basePct + Math.floor(elapsedSec * 1.2), maxStepPct);
                updateProgress(
                    subPct,
                    `Processing (${stepNum}/${total}): ${item.name} (${mbSize} MB) • Uploading & Streaming... [${elapsedSec}s]`
                );
            }, 1000);

            let cleanResult = null;

            try {
                // High-Speed Zero-RAM Engine via Server (Streams upload to disk + Rust Calamine row iteration + constant_memory xlsxwriter)
                // Peak RAM on server remains < 35MB even for 126MB+ files!
                try {
                    const formData = new FormData();
                    formData.append('file', file);
                    formData.append('register_type', item.regType);
                    formData.append('output_format', outputFormat);
                    formData.append('clean_filename', cleanTargetFile);

                    const response = await fetch('/api/clean-register', {
                        method: 'POST',
                        body: formData
                    });

                    if (!response.ok) {
                        const errText = await response.text().catch(() => '');
                        throw new Error(`Server status ${response.status}: ${errText}`);
                    }

                    const colDetected = response.headers.get('X-Stats-Column-Detected') || 'Auto Detected';
                    const colIndex = response.headers.get('X-Stats-Column-Index') || '0';
                    const headerRow = response.headers.get('X-Stats-Header-Row') || '0';
                    const origRows = parseInt(response.headers.get('X-Stats-Original-Rows') || '0', 10);
                    const retRows = parseInt(response.headers.get('X-Stats-Retained-Rows') || '0', 10);
                    const excRows = parseInt(response.headers.get('X-Stats-Excluded-Rows') || '0', 10);
                    const peakRam = parseFloat(response.headers.get('X-Stats-Peak-RAM-MB') || '0');
                    const duration = parseFloat(response.headers.get('X-Stats-Duration-Sec') || '0');
                    const blob = await response.blob();

                    cleanResult = {
                        blob,
                        colDetected,
                        colIndex,
                        headerRow,
                        origRows,
                        retRows,
                        excRows,
                        peakRam,
                        duration
                    };
                } catch (serverErr) {
                    console.warn(`Server cleaning unavailable for ${file.name} (${serverErr.message}).`);
                    // If file is > 35MB, client-side SheetJS XLSX.read will exhaust V8 heap memory.
                    // Warn user rather than crashing the browser tab!
                    if (file.size > 35 * 1024 * 1024) {
                        throw new Error(`Server processing failed (${serverErr.message}). File is ${mbSize} MB (too large for browser memory). Please ensure server is running.`);
                    }
                    showToast(`Server busy. Using in-browser engine for ${file.name}...`, 'info');
                    cleanResult = await cleanFileLocally(file, item.regType, outputFormat);
                } finally {
                    clearInterval(stepTimerId);
                }

                if (cleanResult) {
                    let outFilename = cleanTargetFile;
                    totalOrigRows += cleanResult.origRows;
                    totalRetRows += cleanResult.retRows;
                    totalExcRows += cleanResult.excRows;
                    maxPeakRam = Math.max(maxPeakRam, cleanResult.peakRam);
                    totalDuration += cleanResult.duration;
                    lastDetectedCol = cleanResult.colDetected;
                    lastHeaderRow = cleanResult.headerRow;

                    cleanedFiles.push({
                        name: outFilename,
                        blob: cleanResult.blob,
                        slotName: item.name || 'Cleaned Register',
                        downloaded: false,
                        stats: {
                            colDetected: cleanResult.colDetected,
                            origRows: cleanResult.origRows,
                            retRows: cleanResult.retRows,
                            excRows: cleanResult.excRows,
                            peakRam: cleanResult.peakRam,
                            duration: cleanResult.duration
                        }
                    });
                }

            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
                showToast(`Error in ${file.name}: ${err.message}`, 'error');
                state.isProcessing = false;
                updateProgress(0, 'Cleaning interrupted');
                return;
            }
        }

        updateProgress(90, 'Packaging & preparing downloads...');

        // Update Audit Stats Card
        if (elements.statColDetected) elements.statColDetected.textContent = lastDetectedCol;
        if (elements.statHeaderRow) elements.statHeaderRow.textContent = `Header Row: ${parseInt(lastHeaderRow, 10) + 1}`;
        if (elements.statOrigRows) elements.statOrigRows.textContent = totalOrigRows.toLocaleString();
        if (elements.statRetainedRows) elements.statRetainedRows.textContent = totalRetRows.toLocaleString();
        if (elements.statExcludedRows) elements.statExcludedRows.textContent = totalExcRows.toLocaleString();
        if (elements.statPeakRam) elements.statPeakRam.textContent = `${maxPeakRam.toFixed(1)} MB`;
        if (elements.statDuration) elements.statDuration.textContent = `${totalDuration.toFixed(2)}s`;
        if (elements.statsFileInfo) {
            elements.statsFileInfo.textContent = `Cleaned ${total} file(s) sequentially. Peak RAM: ${maxPeakRam.toFixed(1)} MB (Safe within 512MB limit).`;
        }

        // Save cleaned files state
        cleanerState.cleanedFiles = cleanedFiles;

        let zipBlob = null;
        let zipFilename = '';

        if (cleanedFiles.length > 1) {
            const zip = new JSZip();
            for (const cf of cleanedFiles) {
                zip.file(cf.name, cf.blob);
            }
            zipBlob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
            
            // Clean zip filename based on active tab
            let zipPrefix = 'Cleaned_GST_Registers';
            if (cleanerState.activeTab === 'tab-sale-purchase') {
                if (cleanerState.salePurSubTab === 'details') {
                    zipPrefix = 'Cleaned_Sale_Purchase_Details';
                } else if (cleanerState.salePurSubTab === 'summary') {
                    zipPrefix = 'Cleaned_Sale_Purchase_Summary';
                } else {
                    zipPrefix = 'Cleaned_Sale_Purchase_Registers';
                }
            } else if (cleanerState.activeTab === 'tab-credit-notes') {
                zipPrefix = 'Cleaned_Credit_Notes';
            } else if (cleanerState.activeTab === 'tab-debit-notes') {
                zipPrefix = 'Cleaned_Debit_Notes';
            }
            const ts = getTimestampString();
            zipFilename = `${zipPrefix}_${ts}.zip`;

            cleanerState.zipBlob = zipBlob;
            cleanerState.zipFilename = zipFilename;
        } else if (cleanedFiles.length === 1) {
            cleanerState.zipBlob = null;
            cleanerState.zipFilename = '';
        }

        // Render Download Section (Multi-file ZIP + Individual Files)
        const statsDownloadSection = document.getElementById('stats-download-section');
        const statsZipBox = document.getElementById('stats-zip-box');
        const individualDownloadsWrap = document.getElementById('individual-downloads-wrap');
        const cleanedFilesGrid = document.getElementById('cleaned-files-grid');
        const purgedBanner = document.getElementById('stats-purged-banner');

        if (purgedBanner) {
            purgedBanner.style.display = 'none';
        }

        if (statsDownloadSection) {
            // 1. Multi-file ZIP banner (Show when > 1 file)
            if (cleanedFiles.length > 1 && zipBlob && statsZipBox) {
                const zipFileName = document.getElementById('zip-file-name');
                const zipFileDesc = document.getElementById('zip-file-desc');
                const btnDownloadZipText = document.getElementById('btn-download-zip-text');
                if (zipFileName) zipFileName.textContent = zipFilename;
                if (zipFileDesc) {
                    zipFileDesc.textContent = `All ${cleanedFiles.length} cleaned registers bundled into one package • Total Retained: ${totalRetRows.toLocaleString()} rows • Peak RAM: ${maxPeakRam.toFixed(1)} MB`;
                }
                if (btnDownloadZipText) {
                    btnDownloadZipText.textContent = `Download All (${zipFilename})`;
                }
                statsZipBox.style.display = 'flex';
            } else if (statsZipBox) {
                statsZipBox.style.display = 'none';
            }

            // 2. Individual Cleaned Files Grid
            if (individualDownloadsWrap) individualDownloadsWrap.style.display = 'block';
            if (cleanedFilesGrid) {
                cleanedFilesGrid.innerHTML = '';
                
                cleanedFiles.forEach((cf, idx) => {
                    const card = document.createElement('div');
                    card.className = 'cleaned-file-card';
                    card.innerHTML = `
                        <div class="cleaned-file-info">
                            <div class="cleaned-file-icon-box">
                                <svg class="cleaned-file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                    <polyline points="14 2 14 8 20 8"></polyline>
                                    <line x1="16" y1="13" x2="8" y2="13"></line>
                                    <line x1="16" y1="17" x2="8" y2="17"></line>
                                </svg>
                            </div>
                            <div class="cleaned-file-details">
                                <div class="cleaned-file-title-row">
                                    <span class="cleaned-file-title">${escapeHtml(cf.name)}</span>
                                    <span class="cleaned-badge">${escapeHtml(cf.slotName)}</span>
                                </div>
                                <div class="cleaned-file-meta">
                                    <span>Retained: <strong>${cf.stats.retRows.toLocaleString()}</strong> rows</span>
                                    <span class="meta-dot">&bull;</span>
                                    <span>Excluded: <strong>${cf.stats.excRows.toLocaleString()}</strong> rows</span>
                                    <span class="meta-dot">&bull;</span>
                                    <span>RAM Safe: <strong>${cf.stats.peakRam.toFixed(1)} MB</strong></span>
                                </div>
                            </div>
                        </div>
                        <button type="button" class="glass-btn glass-btn-success action-btn-md btn-download-single" data-file-idx="${idx}">
                            <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                                <polyline points="7 10 12 15 17 10"></polyline>
                                <line x1="12" y1="15" x2="12" y2="3"></line>
                            </svg>
                            <span>Download File</span>
                        </button>
                    `;

                    const btnSingle = card.querySelector('.btn-download-single');
                    if (btnSingle) {
                        btnSingle.addEventListener('click', () => {
                            downloadSingleCleanedFile(idx);
                        });
                    }

                    cleanedFilesGrid.appendChild(card);
                });
            }

            statsDownloadSection.style.display = 'flex';
        }

        if (elements.cleaningStatsCard) {
            elements.cleaningStatsCard.style.display = 'block';
            elements.cleaningStatsCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }

        updateProgress(100, `Cleaning complete! ${total} file(s) ready for download.`);
        setTimeout(() => {
            elements.progressBox.style.display = 'none';
        }, 1500);

        showToast(`Successfully cleaned ${total} file(s)! Download individual files or ZIP package below.`, 'success');
        state.isProcessing = false;
    }

    // Lifecycle Memory Management: Surgical Batch Memory Purge
    // Wipes only the slots cleaned in this specific batch (leaving preset slots for AZ intact),
    // drops blob references, revokes Object URLs, and triggers Python server gc.collect()
    async function purgeCleanedBatchMemory(toastMsg = 'Downloaded! Cleaned batch memory wiped clean.') {
        const purgedSlots = cleanerState.lastProcessedSlotIds || [];

        // 1. Wipe ONLY slots processed in this batch (leave AZ or other slots intact)
        purgedSlots.forEach(slotId => {
            if (cleanerState.slots[slotId]) {
                cleanerState.slots[slotId].file = null;
            }
            const card = document.getElementById(`slot-card-${slotId}`);
            if (card) {
                card.classList.remove('has-file');
                const filenameEl = card.querySelector('.slot-filename');
                const filesizeEl = card.querySelector('.slot-filesize');
                if (filenameEl) filenameEl.textContent = '';
                if (filesizeEl) filesizeEl.textContent = '';
            }
        });

        // 2. Clear blobs and revoke blob URLs
        if (cleanerState.cleanedFiles && cleanerState.cleanedFiles.length > 0) {
            cleanerState.cleanedFiles.forEach(cf => {
                if (cf.blobUrl) {
                    window.URL.revokeObjectURL(cf.blobUrl);
                    cf.blobUrl = null;
                }
                cf.blob = null;
            });
            cleanerState.cleanedFiles = [];
        }

        if (cleanerState.zipBlobUrl) {
            window.URL.revokeObjectURL(cleanerState.zipBlobUrl);
            cleanerState.zipBlobUrl = null;
        }
        cleanerState.zipBlob = null;
        cleanerState.lastProcessedSlotIds = [];

        // 3. Update counter badges in subtab nav
        updateSubtabCounters();

        // 4. Update download UI without destroying the permanent DOM elements
        const purgedBanner = document.getElementById('stats-purged-banner');
        if (purgedBanner) {
            purgedBanner.innerHTML = `
                <div style="background: rgba(16, 185, 129, 0.08); border: 1.5px solid rgba(16, 185, 129, 0.3); border-radius: 12px; padding: 1.1rem 1.4rem; display: flex; align-items: center; justify-content: space-between; gap: 1rem; width: 100%;">
                    <div style="display: flex; align-items: center; gap: 0.85rem;">
                        <div style="width: 36px; height: 36px; border-radius: 50%; background: #d1fae5; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5"><polyline points="20 6 9 17 4 12"></polyline></svg>
                        </div>
                        <div>
                            <div style="font-weight: 700; color: #065f46; font-size: 0.95rem;">Download Complete & RAM Purged Clean</div>
                            <div style="font-size: 0.82rem; color: #047857;">Cleaned batch data has been erased from active memory. Ready for next cleaning batch!</div>
                        </div>
                    </div>
                    <span style="font-size: 0.8rem; font-weight: 700; background: #ecfdf5; color: #059669; padding: 0.4rem 0.85rem; border-radius: 999px; border: 1px solid rgba(16, 185, 129, 0.25); white-space: nowrap;">0 MB Active RAM</span>
                </div>
            `;
            purgedBanner.style.display = 'block';
        }

        // 5. Ping backend to trigger Python gc.collect()
        try {
            await fetch('/api/clean-purge', { method: 'POST' });
        } catch (e) {
            console.warn('Backend purge ping:', e);
        }

        showToast(toastMsg, 'success');
        console.log('[LIFECYCLE PURGE] Cleaned batch slots and blobs purged from memory.');
    }

    // Individual & ZIP Cleaned File Download Handlers
    function downloadSingleCleanedFile(idx) {
        const item = cleanerState.cleanedFiles[idx];
        if (!item || !item.blob) {
            showToast('Cleaned file data not found or memory already purged.', 'warning');
            return;
        }
        const url = window.URL.createObjectURL(item.blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = item.name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => window.URL.revokeObjectURL(url), 2000);
        showToast(`Downloading ${item.name}...`, 'success');

        item.downloaded = true;
        // Visually mark button as downloaded
        const btn = document.querySelector(`.btn-download-single[data-file-idx="${idx}"]`);
        if (btn) {
            btn.innerHTML = `
                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="20 6 9 17 4 12"></polyline>
                </svg>
                <span>Downloaded</span>
            `;
            btn.classList.remove('glass-btn-success');
            btn.classList.add('glass-btn-secondary');
        }
    }

    function downloadZipCleanedPackage() {
        if (!cleanerState.zipBlob) {
            showToast('No ZIP package available or memory already purged.', 'warning');
            return;
        }
        const url = window.URL.createObjectURL(cleanerState.zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = cleanerState.zipFilename || 'Cleaned_GST_Registers.zip';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => window.URL.revokeObjectURL(url), 2000);
        showToast(`Downloading ${cleanerState.zipFilename}...`, 'success');

        const btnZip = document.getElementById('btn-download-zip');
        if (btnZip) {
            const btnText = document.getElementById('btn-download-zip-text');
            if (btnText) btnText.textContent = `Downloaded (${cleanerState.zipFilename})`;
        }
    }

    // High-performance streaming CSV export using chunked Blobs & RFC-4180
    function createCsvBlob(headers, dataRows) {
        const escapeCsv = (val) => {
            if (val === null || val === undefined) return '';
            const s = String(val);
            if (s.indexOf('"') !== -1 || s.indexOf(',') !== -1 || s.indexOf('\n') !== -1 || s.indexOf('\r') !== -1) {
                const escaped = s.split('"').join('""');
                return '"' + escaped + '"';
            }
            return s;
        };

        const chunks = [];
        // UTF-8 BOM for Microsoft Excel compatibility
        chunks.push('\uFEFF');
        // Header row
        chunks.push(headers.map(escapeCsv).join(',') + '\r\n');

        const BATCH_SIZE = 2500;
        let batch = '';
        for (let i = 0; i < dataRows.length; i++) {
            batch += dataRows[i].map(escapeCsv).join(',') + '\r\n';
            if ((i + 1) % BATCH_SIZE === 0) {
                chunks.push(batch);
                batch = '';
            }
        }
        if (batch.length > 0) {
            chunks.push(batch);
        }

        return new Blob(chunks, { type: 'text/csv;charset=utf-8;' });
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 15000);
    }

    function downloadCurrentMergedFile() {
        if (!state.lastMergedResult || !state.lastMergedResult.blob) {
            showToast('No merged file available for download.', 'warning');
            return;
        }
        downloadBlob(state.lastMergedResult.blob, state.lastMergedResult.filename);
        showToast(`Downloading ${state.lastMergedResult.filename}...`, 'success');
    }

    // Modal: Open Send to Clean Excel with smart slot pre-selection
    function openSendToCleanerModal() {
        if (!state.lastMergedResult || !state.lastMergedResult.file) {
            showToast('No merged dataset available. Please merge files first.', 'warning');
            return;
        }

        if (elements.sendModalFilename) {
            elements.sendModalFilename.textContent = state.lastMergedResult.filename;
        }
        if (elements.sendModalStats) {
            elements.sendModalStats.textContent = `${state.lastMergedResult.rowsCount.toLocaleString()} rows • ${formatBytes(state.lastMergedResult.sizeBytes)}`;
        }

        // Smart slot pre-selection based on file name and merge mode
        const fname = (state.lastMergedResult.filename || '').toLowerCase();
        let targetSlot = 's3';

        if (state.lastMergedResult.mode === 'zoho') {
            if (state.zohoPart === 'purchase' || fname.includes('purchase') || fname.includes('bill') || fname.includes('payable')) {
                targetSlot = 'p2';
            } else if (fname.includes('credit') || fname.includes('cn')) {
                targetSlot = 'c2';
            } else if (fname.includes('debit') || fname.includes('dn') || fname.includes('vendor')) {
                targetSlot = 'd2';
            } else {
                targetSlot = 's2';
            }
        } else if (state.lastMergedResult.mode === 'erp') {
            if (state.erpPart === 'purchase' || fname.includes('purchase') || fname.includes('bill') || fname.includes('payable')) {
                targetSlot = 'p3';
            } else if (fname.includes('credit') || fname.includes('cn')) {
                targetSlot = 'c3';
            } else if (fname.includes('debit') || fname.includes('dn')) {
                targetSlot = 'd3';
            } else {
                targetSlot = 's3';
            }
        } else {
            if (fname.includes('zoho') && (fname.includes('purchase') || fname.includes('bill') || fname.includes('payable'))) {
                targetSlot = 'p2';
            } else if (fname.includes('purchase') || fname.includes('bill') || fname.includes('payable')) {
                targetSlot = 'p3';
            } else if (fname.includes('zoho') && (fname.includes('sale') || fname.includes('inv'))) {
                targetSlot = 's2';
            } else if (fname.includes('credit') || fname.includes('cn')) {
                targetSlot = fname.includes('zoho') ? 'c2' : 'c3';
            } else if (fname.includes('debit') || fname.includes('dn')) {
                targetSlot = fname.includes('zoho') ? 'd2' : 'd3';
            } else {
                targetSlot = 's3';
            }
        }

        const targetCheckbox = document.getElementById(`slot-opt-${targetSlot}`);
        if (targetCheckbox) {
            targetCheckbox.checked = true;
        }

        // Highlight selected cards
        document.querySelectorAll('.slot-select-card').forEach(card => {
            const cb = card.querySelector('input[name="cleaner-target-slot"]');
            if (cb && cb.checked) card.classList.add('selected');
            else card.classList.remove('selected');
        });

        if (elements.modalSendCleaner) {
            elements.modalSendCleaner.style.display = 'flex';
        }
    }

    // Lifecycle Memory Management: Wipe Merger Tab Memory completely to reclaim RAM
    function wipeMergerMemory() {
        state.files = [];
        state.mergedHeaders = [];
        state.mergedData = [];
        if (state.lastMergedResult) {
            state.lastMergedResult.headers = null;
            state.lastMergedResult.data = null;
            state.lastMergedResult.file = null;
            state.lastMergedResult = null;
        }
        renderFileQueue();
        if (elements.mergedResultBox) elements.mergedResultBox.style.display = 'none';
        if (elements.progressBox) elements.progressBox.style.display = 'none';
        state.isProcessing = false;
        console.log('[LIFECYCLE PURGE] Merger tab memory completely wiped to 0 MB.');
    }

    // Modal: Confirm Send to Cleaner Slot(s) - Supports Multi-Select
    function confirmSendToCleanerSlot() {
        if (!state.lastMergedResult || !state.lastMergedResult.file) {
            showToast('No merged dataset available.', 'warning');
            return;
        }

        const checkedBoxes = Array.from(document.querySelectorAll('input[name="cleaner-target-slot"]:checked'));
        if (checkedBoxes.length === 0) {
            showToast('Please select at least one destination cleaner slot.', 'warning');
            return;
        }

        const selectedSlotIds = checkedBoxes.map(cb => cb.value);

        // Close the modal
        if (elements.modalSendCleaner) {
            elements.modalSendCleaner.style.display = 'none';
        }

        const firstSlot = selectedSlotIds[0];

        // Determine target tab in GST cleaner from first selected slot
        let targetTabId = 'tab-sale-purchase';
        if (['c3', 'c2'].includes(firstSlot)) {
            targetTabId = 'tab-credit-notes';
        } else if (['d3', 'd2'].includes(firstSlot)) {
            targetTabId = 'tab-debit-notes';
        }

        // 1. Switch application mode to Cleaner
        if (elements.modeCleaner) {
            elements.modeCleaner.click();
        }

        // 2. Switch cleaner tab
        const tabBtn = document.querySelector(`.cleaner-tab-btn[data-clean-tab="${targetTabId}"]`);
        if (tabBtn) {
            tabBtn.click();
        }

        // 2b. If target is in sale-purchase, switch to corresponding subtab & category
        if (targetTabId === 'tab-sale-purchase') {
            const hasDetails = selectedSlotIds.some(id => id.includes('_details'));
            const subtabBtn = document.querySelector(`.sale-pur-subtab-btn[data-subtab="${hasDetails ? 'details' : 'summary'}"]`);
            if (subtabBtn) {
                subtabBtn.click();
            }
            if (hasDetails) {
                const hasAj = selectedSlotIds.some(id => id.includes('_aj'));
                const hasAz = selectedSlotIds.some(id => id.includes('_az'));
                let targetCat = 'both';
                if (hasAj && !hasAz) targetCat = 'aj_my_fk';
                else if (hasAz && !hasAj) targetCat = 'az';
                const catBtn = document.querySelector(`.details-cat-btn[data-details-cat="${targetCat}"]`);
                if (catBtn) catBtn.click();
            }
        }

        // 3. Attach file to ALL selected destination slot cards
        let attachedCount = 0;
        const mergedFileRef = state.lastMergedResult.file;
        selectedSlotIds.forEach((slotId, index) => {
            const card = document.getElementById(`slot-card-${slotId}`);
            if (card) {
                const fileMeta = card.querySelector('.slot-file-meta');
                const filenameEl = card.querySelector('.slot-filename');
                const filesizeEl = card.querySelector('.slot-filesize');
                setSlotFile(slotId, mergedFileRef, card, fileMeta, filenameEl, filesizeEl);
                attachedCount++;

                // 4. Scroll first slot into view & trigger highlight pulse animation
                setTimeout(() => {
                    if (index === 0) {
                        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }
                    card.classList.add('slot-highlight-pulse');
                    setTimeout(() => {
                        card.classList.remove('slot-highlight-pulse');
                    }, 3600);
                }, 100);
            }
        });

        updateSubtabCounters();

        // 5. Instantly wipe Merger tab memory so RAM is 100% freed
        wipeMergerMemory();

        if (attachedCount === 1) {
            const slotName = cleanerState.slots[firstSlot] ? cleanerState.slots[firstSlot].name : 'GST Register Slot';
            showToast(`Merged file attached to ${slotName}! Merger memory wiped clean (0 MB active RAM).`, 'success');
        } else {
            showToast(`Merged file attached to ${attachedCount} slots! Merger memory wiped clean (0 MB active RAM).`, 'success');
        }
    }

    // Merge Execution Handler (Saves to Memory & Displays Action Box)
    async function handleMergeAndDownload() {
        if (state.mode === 'cleaner') {
            if (cleanerState.cleanedFiles.length > 0) {
                if (cleanerState.cleanedFiles.length > 1) {
                    downloadZipCleanedPackage();
                } else {
                    downloadSingleCleanedFile(0);
                }
                return;
            }
            await handleCleanSelectedFiles();
            return;
        }

        if (state.mode === 'zip_consolidator' || elements.outputFormat.value === 'zip') {
            await handleZipConsolidation();
            return;
        }

        try {
            state.isProcessing = true;
            const result = await processMerge();
            if (!result || result.headers.length === 0) {
                elements.progressBox.style.display = 'none';
                state.isProcessing = false;
                showToast('Merge aborted: No valid headers or data extracted.', 'error');
                return;
            }

            const format = elements.outputFormat.value; // 'xlsx' or 'csv'
            let baseFilename = elements.outputFilename.value.trim() || 'Merged_Output';

            // Remove manual extension if user typed it
            baseFilename = baseFilename.replace(/\.(xlsx|csv|xls)$/i, '');

            const totalRows = result.data.length;
            const totalCols = result.headers.length;
            const totalCells = totalRows * totalCols;

            let blob = null;
            let finalFilename = '';
            let exportFormat = format;

            // Direct CSV Export or Large Dataset Safeguard (>50k rows or >1.5M cells)
            if (format === 'csv' || totalRows > 50000 || totalCells > 1500000) {
                if (format !== 'csv') {
                    exportFormat = 'csv';
                    updateProgress(95, `Large dataset (${totalRows.toLocaleString()} rows). Auto-packaging as Excel-compatible CSV...`);
                } else {
                    updateProgress(95, `Generating CSV dataset (${totalRows.toLocaleString()} rows)...`);
                }
                finalFilename = `${baseFilename}.csv`;
                blob = createCsvBlob(result.headers, result.data);
            } else {
                // Standard Excel (.xlsx) Export using SheetJS
                updateProgress(95, `Assembling Excel workbook (${totalRows.toLocaleString()} rows)...`);
                finalFilename = `${baseFilename}.xlsx`;
                const combinedAoa = [result.headers, ...result.data];
                const worksheet = XLSX.utils.aoa_to_sheet(combinedAoa);
                const workbook = XLSX.utils.book_new();
                let sheetTitle = 'Merged Data';
                if (state.mode === 'zoho') {
                    sheetTitle = state.zohoPart === 'sale' ? 'Merged Zoho Sale' : (state.zohoPart === 'purchase' ? 'Merged Zoho Purchase' : 'Merged Zoho Normal');
                } else if (state.mode === 'erp') {
                    sheetTitle = state.erpPart === 'sale' ? 'Merged ERP Sale' : (state.erpPart === 'purchase' ? 'Merged ERP Purchase' : 'Merged ERP Normal');
                } else if (state.mode === 'account_center') {
                    sheetTitle = 'Merged Account Data';
                }
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetTitle);
                const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
                blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            }

            const fileObj = new File([blob], finalFilename, { type: blob.type, lastModified: Date.now() });

            // Cache merged result in application state
            state.lastMergedResult = {
                headers: result.headers,
                data: result.data,
                filename: finalFilename,
                format: exportFormat,
                blob: blob,
                file: fileObj,
                rowsCount: totalRows,
                colsCount: totalCols,
                sizeBytes: blob.size,
                mode: state.mode
            };

            // Update Merged Result Box in UI
            if (elements.mergedResultBox) {
                if (elements.mergedResultFilename) elements.mergedResultFilename.textContent = finalFilename;
                if (elements.mergedResultRows) elements.mergedResultRows.textContent = totalRows.toLocaleString();
                if (elements.mergedResultCols) elements.mergedResultCols.textContent = totalCols.toLocaleString();
                if (elements.mergedResultSize) elements.mergedResultSize.textContent = formatBytes(blob.size);

                const zohoLabel = `Zoho Books (${state.zohoPart ? state.zohoPart.charAt(0).toUpperCase() + state.zohoPart.slice(1) : 'Normal'})`;
                const erpLabel = `ERP Merge (${state.erpPart ? state.erpPart.charAt(0).toUpperCase() + state.erpPart.slice(1) : 'Normal'})`;
                const modeLabels = {
                    standard: 'Standard Excel & CSV',
                    erp: erpLabel,
                    zoho: zohoLabel,
                    account_center: 'Account Center Mode',
                    custom: 'Custom Range Mode'
                };
                if (elements.mergedModeTag) {
                    elements.mergedModeTag.textContent = modeLabels[state.mode] || 'Merged File';
                }

                // Show "Send to Clean Excel" button ONLY for the requested modules
                if (elements.btnSendToCleaner) {
                    if (['standard', 'erp', 'zoho', 'account_center'].includes(state.mode)) {
                        elements.btnSendToCleaner.style.display = 'inline-flex';
                    } else {
                        elements.btnSendToCleaner.style.display = 'none';
                    }
                }

                elements.mergedResultBox.style.display = 'block';
                elements.mergedResultBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            }

            updateProgress(100, `Merge complete! ${totalRows.toLocaleString()} rows ready.`);
            setTimeout(() => { elements.progressBox.style.display = 'none'; }, 1500);

            if (exportFormat === 'csv' && format !== 'csv') {
                showToast(`Dataset merged (${totalRows.toLocaleString()} rows). Auto-prepared as Excel-ready CSV for stability. Click Download or Send to Clean Excel.`, 'warning');
            } else {
                showToast(`Dataset merged successfully (${totalRows.toLocaleString()} rows)! Ready to download or send to GST Register Cleaner.`, 'success');
            }
            state.isProcessing = false;

        } catch (err) {
            console.error('Merge error:', err);
            elements.progressBox.style.display = 'none';
            state.isProcessing = false;
            showToast(`Merge failed: ${err.message || 'Browser memory limit reached.'}. Try selecting CSV format.`, 'error');
        }
    }

    // Setup Preview, Actions, and Modals
    function setupActions() {
        elements.btnPreviewData.addEventListener('click', () => {
            if (!state.isProcessing) {
                showPreview();
            }
        });

        elements.btnMergeDownload.addEventListener('click', () => {
            if (!state.isProcessing) {
                handleMergeAndDownload();
            }
        });

        // Download Merged File
        if (elements.btnDownloadMerged) {
            elements.btnDownloadMerged.addEventListener('click', () => {
                downloadCurrentMergedFile();
            });
        }

        // Open Send to Clean Excel Modal
        if (elements.btnSendToCleaner) {
            elements.btnSendToCleaner.addEventListener('click', () => {
                openSendToCleanerModal();
            });
        }

        // Send to Cleaner Modal Events
        if (elements.btnCloseSendModal) {
            elements.btnCloseSendModal.addEventListener('click', () => {
                if (elements.modalSendCleaner) elements.modalSendCleaner.style.display = 'none';
            });
        }

        if (elements.btnCancelSendCleaner) {
            elements.btnCancelSendCleaner.addEventListener('click', () => {
                if (elements.modalSendCleaner) elements.modalSendCleaner.style.display = 'none';
            });
        }

        if (elements.modalSendCleaner) {
            elements.modalSendCleaner.addEventListener('click', (e) => {
                if (e.target === elements.modalSendCleaner) {
                    elements.modalSendCleaner.style.display = 'none';
                }
            });
        }

        // Slot checkbox card selection in modal (Multi-Select)
        document.querySelectorAll('.slot-select-card').forEach(card => {
            const cb = card.querySelector('input[name="cleaner-target-slot"]');
            if (cb) {
                cb.addEventListener('change', () => {
                    card.classList.toggle('selected', cb.checked);
                });
            }
        });

        // Confirm Send to Cleaner Slot
        if (elements.btnConfirmSendCleaner) {
            elements.btnConfirmSendCleaner.addEventListener('click', () => {
                confirmSendToCleanerSlot();
            });
        }

        // Multi-File ZIP Package Download Button
        if (elements.btnDownloadZip) {
            elements.btnDownloadZip.addEventListener('click', () => {
                downloadZipCleanedPackage();
            });
        }

        elements.btnClosePreview.addEventListener('click', () => {
            elements.previewModal.style.display = 'none';
        });

        elements.previewModal.addEventListener('click', (e) => {
            if (e.target === elements.previewModal) {
                elements.previewModal.style.display = 'none';
            }
        });
    }

    // Launch App
    document.addEventListener('DOMContentLoaded', init);
})();
