/**
 * Main JavaScript for PDF translation service with batch processing.
 */

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('upload-form');
    const fileInput = document.getElementById('file-input');
    const uploadArea = document.getElementById('upload-area');
    const fileQueue = document.getElementById('file-queue');
    const queueList = document.getElementById('queue-list');
    const clearAllBtn = document.getElementById('clear-all');
    const submitBtn = document.getElementById('submit-btn');
    const stopBtn = document.getElementById('stop-btn');
    const batchProgress = document.getElementById('batch-progress');
    const batchProgressText = document.getElementById('batch-progress-text');
    const batchProgressFill = document.getElementById('batch-progress-fill');
    const resultsSection = document.getElementById('results-section');
    const resultsList = document.getElementById('results-list');
    const errorSection = document.getElementById('error-section');
    const errorText = document.getElementById('error-text');

    // File queue: array of { id, file, status, result, abortController }
    let fileQueueData = [];
    let queueIdCounter = 0;
    let isProcessing = false;
    const MAX_CONCURRENT = 2; // Process 2 files at a time

    // Click to select files
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // Drag and drop
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        handleFilesSelect(e.dataTransfer.files);
    });

    // File input change
    fileInput.addEventListener('change', (e) => {
        handleFilesSelect(e.target.files);
        fileInput.value = ''; // Reset to allow re-selecting same files
    });

    // Clear all files
    clearAllBtn.addEventListener('click', () => {
        if (!isProcessing) {
            clearAllFiles();
        }
    });

    // Stop button
    stopBtn.addEventListener('click', () => {
        isProcessing = false;
        // Abort all processing files
        for (const item of fileQueueData) {
            if (item.status === 'processing' && item.abortController) {
                item.abortController.abort();
                item.status = 'error';
                item.error = '已取消';
                updateQueueItemStatus(item.id, 'error');
            }
        }
        hideStopButton();
        submitBtn.disabled = false;
        updateUI();
        showResults();
    });

    // Form submit
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (fileQueueData.length === 0 || isProcessing) return;
        await processBatch();
    });

    /**
     * Handle multiple file selection.
     */
    function handleFilesSelect(files) {
        const validFiles = [];
        const maxSize = 20 * 1024 * 1024;

        for (const file of files) {
            if (!file.name.toLowerCase().endsWith('.pdf')) {
                continue; // Skip non-PDF files silently
            }
            if (file.size > maxSize) {
                showError(`${file.name} 超過 20MB 限制`);
                continue;
            }
            // Check for duplicates
            const exists = fileQueueData.some(f => f.file.name === file.name && f.file.size === file.size);
            if (!exists) {
                validFiles.push(file);
            }
        }

        if (validFiles.length === 0) return;

        // Add files to queue
        for (const file of validFiles) {
            const queueItem = {
                id: ++queueIdCounter,
                file: file,
                status: 'pending', // pending, processing, completed, error
                result: null,
                error: null
            };
            fileQueueData.push(queueItem);
            addQueueItemToDOM(queueItem);
        }

        updateUI();
        hideError();
    }

    /**
     * Add a queue item to the DOM.
     */
    function addQueueItemToDOM(item) {
        const li = document.createElement('li');
        li.className = 'queue-item pending';
        li.id = `queue-item-${item.id}`;
        li.innerHTML = `
            <span class="queue-item-name">${item.file.name}</span>
            <div class="queue-item-status"></div>
            <button type="button" class="btn-remove" data-id="${item.id}">✕</button>
        `;

        // Add remove button handler
        li.querySelector('.btn-remove').addEventListener('click', (e) => {
            if (!isProcessing) {
                removeQueueItem(parseInt(e.target.dataset.id));
            }
        });

        queueList.appendChild(li);
    }

    /**
     * Remove a queue item.
     */
    function removeQueueItem(id) {
        fileQueueData = fileQueueData.filter(f => f.id !== id);
        const li = document.getElementById(`queue-item-${id}`);
        if (li) li.remove();
        updateUI();
    }

    /**
     * Clear all files from queue.
     */
    function clearAllFiles() {
        fileQueueData = [];
        queueList.innerHTML = '';
        updateUI();
        hideResults();
    }

    /**
     * Update UI based on queue state.
     */
    function updateUI() {
        const hasFiles = fileQueueData.length > 0;
        const hasPendingFiles = fileQueueData.some(f => f.status === 'pending');

        fileQueue.classList.toggle('hidden', !hasFiles);
        submitBtn.disabled = !hasPendingFiles || isProcessing;

        // Hide remove buttons during processing
        const removeButtons = queueList.querySelectorAll('.btn-remove');
        removeButtons.forEach(btn => {
            btn.style.display = isProcessing ? 'none' : 'block';
        });
    }

    /**
     * Update queue item status in DOM.
     */
    function updateQueueItemStatus(id, status, progressText = '') {
        const li = document.getElementById(`queue-item-${id}`);
        if (!li) return;

        li.className = `queue-item ${status}`;
        const statusDiv = li.querySelector('.queue-item-status');

        if (status === 'processing') {
            statusDiv.innerHTML = `<div class="spinner-small"></div><span class="progress-text">${progressText || '準備中...'}</span>`;
        } else {
            statusDiv.innerHTML = '';
        }
    }

    /**
     * Process all files in the queue with parallel processing.
     */
    async function processBatch() {
        isProcessing = true;
        submitBtn.disabled = true;
        showStopButton();
        hideError();
        hideResults();

        const pendingFiles = fileQueueData.filter(f => f.status === 'pending');
        const totalFiles = pendingFiles.length;
        let completedFiles = 0;

        // Show batch progress
        batchProgress.classList.remove('hidden');
        updateBatchProgress(0, totalFiles);

        // Process a single file and return when done
        async function processFile(item) {
            if (!isProcessing) return;

            item.status = 'processing';
            updateQueueItemStatus(item.id, 'processing');

            try {
                const result = await uploadAndTranslateFile(item);
                item.status = 'completed';
                item.result = result;
                updateQueueItemStatus(item.id, 'completed');
            } catch (error) {
                item.status = 'error';
                item.error = error.message || '翻譯失敗';
                updateQueueItemStatus(item.id, 'error');
            }

            completedFiles++;
            updateBatchProgress(completedFiles, totalFiles);
        }

        // Process files with concurrency limit
        const queue = [...pendingFiles];
        const activePromises = new Set();

        while (queue.length > 0 || activePromises.size > 0) {
            if (!isProcessing) break;

            // Start new tasks up to MAX_CONCURRENT
            while (queue.length > 0 && activePromises.size < MAX_CONCURRENT && isProcessing) {
                const item = queue.shift();
                const promise = processFile(item).then(() => {
                    activePromises.delete(promise);
                });
                activePromises.add(promise);
            }

            // Wait for at least one to complete
            if (activePromises.size > 0) {
                await Promise.race(activePromises);
            }
        }

        // Done processing
        isProcessing = false;
        hideStopButton();
        submitBtn.disabled = !fileQueueData.some(f => f.status === 'pending');
        updateUI();

        // Show results
        showResults();
    }

    /**
     * Upload and translate a single file.
     */
    async function uploadAndTranslateFile(item) {
        return new Promise(async (resolve, reject) => {
            const formData = new FormData();
            formData.append('file', item.file);

            item.abortController = new AbortController();

            try {
                const response = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData,
                    signal: item.abortController.signal
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.detail || '上傳失敗');
                }

                // Handle SSE stream
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                let buffer = '';
                let result = null;

                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;

                    buffer += decoder.decode(value, { stream: true });

                    const events = buffer.split('\n\n');
                    buffer = events.pop();

                    for (const eventStr of events) {
                        if (!eventStr.trim()) continue;

                        const lines = eventStr.split('\n');
                        let eventType = '';
                        let eventData = '';

                        for (const line of lines) {
                            if (line.startsWith('event: ')) {
                                eventType = line.slice(7);
                            } else if (line.startsWith('data: ')) {
                                eventData = line.slice(6);
                            }
                        }

                        if (eventType && eventData) {
                            const data = JSON.parse(eventData);
                            if (eventType === 'progress') {
                                // Update individual file progress
                                updateQueueItemStatus(item.id, 'processing', data.message);
                            } else if (eventType === 'complete') {
                                result = data;
                            } else if (eventType === 'error') {
                                throw new Error(data.detail || '翻譯失敗');
                            }
                        }
                    }
                }

                if (result) {
                    resolve(result);
                } else {
                    reject(new Error('未收到翻譯結果'));
                }

            } catch (error) {
                if (error.name === 'AbortError') {
                    reject(new Error('已取消'));
                } else {
                    reject(error);
                }
            } finally {
                item.abortController = null;
            }
        });
    }

    /**
     * Update batch progress bar.
     */
    function updateBatchProgress(completed, total) {
        const percent = total > 0 ? (completed / total) * 100 : 0;
        batchProgressText.textContent = `${completed} / ${total}`;
        batchProgressFill.style.width = `${percent}%`;
    }

    /**
     * Show stop button.
     */
    function showStopButton() {
        stopBtn.classList.remove('hidden');
    }

    /**
     * Hide stop button.
     */
    function hideStopButton() {
        stopBtn.classList.add('hidden');
    }

    /**
     * Show results section.
     */
    function showResults() {
        const completedItems = fileQueueData.filter(f => f.status === 'completed' || f.status === 'error');
        if (completedItems.length === 0) return;

        resultsList.innerHTML = '';

        for (const item of completedItems) {
            const li = document.createElement('li');

            if (item.status === 'completed' && item.result) {
                const originalName = item.result.original_name || 'file';
                const downloadUrl = `${item.result.download_url}?filename=${encodeURIComponent(originalName)}`;

                // Debug download URLs
                const debugConvertedUrl = item.result.debug_converted_url
                    ? `${item.result.debug_converted_url}?filename=${encodeURIComponent(originalName)}`
                    : null;
                const debugFirstPassUrl = item.result.debug_first_pass_url
                    ? `${item.result.debug_first_pass_url}?filename=${encodeURIComponent(originalName)}`
                    : null;

                const stats = item.result.stats;
                const statsText = `${formatTime(stats.processing_time_seconds)} | ${formatBytes(stats.output_file_size_bytes)} | $${stats.estimated_cost_usd.toFixed(4)}`;

                // Build debug buttons HTML
                let debugButtons = '';
                if (debugConvertedUrl || debugFirstPassUrl) {
                    debugButtons = '<div class="debug-buttons">';
                    if (debugConvertedUrl) {
                        debugButtons += `<a href="${debugConvertedUrl}" class="btn-debug" title="PDF轉Word後（翻譯前）">轉檔</a>`;
                    }
                    if (debugFirstPassUrl) {
                        debugButtons += `<a href="${debugFirstPassUrl}" class="btn-debug" title="第一次翻譯後（補翻前）">Pass1</a>`;
                    }
                    debugButtons += '</div>';
                }

                li.className = 'result-item';
                li.innerHTML = `
                    <div class="result-item-info">
                        <div class="result-item-name">${item.file.name}</div>
                        <div class="result-item-stats">${statsText}</div>
                    </div>
                    <div class="result-item-actions">
                        ${debugButtons}
                        <a href="${downloadUrl}" class="btn-download-small">下載</a>
                    </div>
                `;
            } else {
                li.className = 'result-item error';
                li.innerHTML = `
                    <div class="result-item-info">
                        <div class="result-item-name">${item.file.name}</div>
                        <div class="result-item-stats">${item.error || '翻譯失敗'}</div>
                    </div>
                `;
            }

            resultsList.appendChild(li);
        }

        resultsSection.classList.remove('hidden');
    }

    /**
     * Hide results section.
     */
    function hideResults() {
        resultsSection.classList.add('hidden');
        resultsList.innerHTML = '';
    }

    /**
     * Format bytes to human readable.
     */
    function formatBytes(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * Format seconds to human readable.
     */
    function formatTime(seconds) {
        if (seconds < 60) {
            return seconds.toFixed(1) + ' 秒';
        }
        const mins = Math.floor(seconds / 60);
        const secs = (seconds % 60).toFixed(0);
        return `${mins} 分 ${secs} 秒`;
    }

    /**
     * Show error message.
     */
    function showError(message) {
        errorText.textContent = message;
        errorSection.classList.remove('hidden');
    }

    /**
     * Hide error message.
     */
    function hideError() {
        errorSection.classList.add('hidden');
    }
});
