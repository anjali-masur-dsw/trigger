let mailboxItem = null;
let filename = '';
const BACKEND_URL = "https://metamathematical-mariano-interresponsible.ngrok-free.dev";

// UI Elements
let loadingContainer, formContainer, loadingText, loadingSubtext, progressBar, progressContainer;
let statusItems = {};

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        mailboxItem = Office.context.mailbox.item;

        // Initialize UI elements
        loadingContainer = document.getElementById('loadingContainer');
        formContainer = document.getElementById('formContainer');
        loadingText = document.querySelector('.loading-text');
        loadingSubtext = document.querySelector('.loading-subtext');
        progressBar = document.getElementById('progressBar');
        progressContainer = document.getElementById('progressContainer');

        // Map status items
        ['email', 'flow', 'extract', 'render'].forEach(id => {
            statusItems[id] = document.getElementById(`status-${id}`);
        });

        console.log('Office.js initialized successfully');
        triggerFlowAndLoadForm();
    }
});

function updateProgress(percent, statusId, text, subtext) {
    if (progressBar) progressBar.style.width = `${percent}%`;
    if (progressContainer) progressContainer.style.display = 'block';
    if (loadingText) loadingText.textContent = text;
    if (loadingSubtext) loadingSubtext.textContent = subtext;

    if (statusId && statusItems[statusId]) {
        // Reset all
        Object.values(statusItems).forEach(item => item.classList.remove('active'));

        // Set active
        statusItems[statusId].classList.add('active');

        // Set previous as completed
        const order = ['email', 'flow', 'extract', 'render'];
        const currentIndex = order.indexOf(statusId);
        for (let i = 0; i < currentIndex; i++) {
            statusItems[order[i]].classList.add('completed');
            statusItems[order[i]].classList.remove('active');
        }
    }
}

async function triggerFlowAndLoadForm() {
    try {
        updateProgress(5, 'email', 'Initialising...', 'Connecting to underwriter services');

        // Step 1: Get email data
        updateProgress(15, 'email', 'Collecting email data...', 'Analysing message content and attachments');
        const emailData = await getEmailData();
        console.log('Email data collected');

        // Step 2: Trigger Power Automate flow
        updateProgress(30, 'flow', 'Triggering workflow...', 'Sending data to Power Automate for processing');

        const flowUrl = "https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA";

        const response = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(emailData)
        });

        if (!response.ok) {
            throw new Error(`Flow Error: ${response.status}`);
        }

        console.log('Flow triggered successfully');

        // Step 3: Poll for extracted form data
        updateProgress(45, 'extract', 'Extracting form data...', 'Reviewing attachments (this may take a minute)');

        let extractedData = null;
        let pollingAttempts = 0;
        const maxPollingAttempts = 60; // 5 minutes max

        while (pollingAttempts < maxPollingAttempts) {
            try {
                // IMPORTANT: Use the bypass header for ngrok
                const pendingResponse = await fetch(`${BACKEND_URL}/api/pending`, {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });

                if (pendingResponse.ok) {
                    const data = await pendingResponse.json();
                    if (data.success && data.files && data.files.length > 0) {
                        extractedData = data.files[0];
                        break;
                    }
                } else if (pendingResponse.status === 404) {
                    // This is expected while processing
                    console.log("Still waiting for server processing...");
                } else {
                    console.warn(`Server responded with ${pendingResponse.status}`);
                }
            } catch (pollError) {
                console.warn("Polling attempt failed:", pollError.message);
            }

            // Update progress slightly during polling to show life
            const progressRange = 40; // from 45% to 85%
            const currentProgress = 45 + Math.min((pollingAttempts / 20) * progressRange, progressRange);
            updateProgress(currentProgress, 'extract', 'Extracting form data...', `Still working on it... (Attempt ${pollingAttempts + 1})`);

            await new Promise(resolve => setTimeout(resolve, 5000));
            pollingAttempts++;
        }

        if (!extractedData) {
            throw new Error('Timeout: Extraction took too long. Please try refreshing.');
        }

        // Step 4: Populate form
        updateProgress(90, 'render', 'Loading form...', 'Finalising policy details');
        populateForm(extractedData);

        // Success!
        updateProgress(100, 'render', 'Ready!', 'Form loaded successfully');

        setTimeout(() => {
            loadingContainer.style.display = 'none';
            formContainer.style.display = 'block';
        }, 500);

    } catch (error) {
        console.error('Taskpane Error:', error);
        loadingText.textContent = 'Something went wrong';
        loadingSubtext.innerHTML = `<div class="error-message" style="margin: 10px 0;">${error.message}</div>
            <button class="submit-button" onclick="location.reload()" style="margin-top: 10px; padding: 8px 15px; font-size: 12px;">Try Again</button>`;
    }
}

function populateForm(extractedData) {
    filename = extractedData.filename || '';
    const data = extractedData.email_fields || extractedData.extracted_data || {};

    if (data.broker_email) document.getElementById('brokerEmail').value = data.broker_email;
    if (data.broker_name) document.getElementById('brokerName').value = data.broker_name;
    if (data.underwriter_email) document.getElementById('underwriterEmail').value = data.underwriter_email;
    if (data.underwriter_name) document.getElementById('underwriterName').value = data.underwriter_name;
    if (data.policy_number) document.getElementById('policyNumber').value = data.policy_number;
    if (data.broker_agency_name) document.getElementById('agencyName').value = data.broker_agency_name;
    if (data.broker_agency_id || data.agency_id) document.getElementById('agencyId').value = data.broker_agency_id || data.agency_id;
    if (data.email_summary) document.getElementById('emailSummary').value = data.email_summary;
    if (data.comments) document.getElementById('comments').value = data.comments;

    const timestampField = document.getElementById('timestamp');
    const now = new Date();
    timestampField.value = now.toLocaleString();
}

async function getEmailData() {
    const item = mailboxItem;
    const subject = await getSubject(item);
    const body = await getBody(item);
    const attachments = await getAttachmentContents(item);

    return {
        triggeredAt: new Date().toISOString(),
        userEmail: Office.context.mailbox.userProfile.emailAddress || "",
        subject: subject,
        body: body,
        from: getSender(item),
        receivedDateTime: item.dateTimeCreated || new Date().toISOString(),
        itemId: item.itemId || "",
        conversationId: item.conversationId || "",
        attachments: attachments
    };
}

function getSubject(item) {
    return new Promise(resolve => {
        if (typeof item.subject === "string") resolve(item.subject);
        else item.subject.getAsync(r => resolve(r.status === "succeeded" ? r.value : ""));
    });
}

function getBody(item) {
    return new Promise(resolve => {
        item.body.getAsync(Office.CoercionType.Text, r => resolve(r.status === "succeeded" ? r.value : ""));
    });
}

function getSender(item) {
    if (!item.from) return "Unknown";
    return item.from.emailAddress || item.from.displayName || "Unknown";
}

async function getAttachmentContents(item) {
    if (!item.attachments) return [];
    const results = [];
    for (const att of item.attachments) {
        try {
            await new Promise(resolve => {
                item.getAttachmentContentAsync(att.id, res => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        results.push({
                            id: att.id,
                            name: att.name,
                            contentType: att.contentType,
                            content: res.value.content,
                            format: res.value.format
                        });
                    }
                    resolve();
                });
            });
        } catch (e) { console.error(e); }
    }
    return results;
}

// Handle form submission
document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('insuranceForm');
    if (form) form.addEventListener('submit', handleFormSubmit);
});

async function handleFormSubmit(e) {
    e.preventDefault();
    const submitButton = document.querySelector('.submit-button');
    const successMessage = document.getElementById('successMessage');

    const formData = {
        broker_email: document.getElementById('brokerEmail').value,
        broker_name: document.getElementById('brokerName').value,
        underwriter_email: document.getElementById('underwriterEmail').value,
        underwriter_name: document.getElementById('underwriterName').value,
        policy_number: document.getElementById('policyNumber').value,
        broker_agency_name: document.getElementById('agencyName').value,
        broker_agency_id: document.getElementById('agencyId').value,
        email_summary: document.getElementById('emailSummary').value,
        comments: document.getElementById('comments').value,
        timestamp: document.getElementById('timestamp').value
    };

    try {
        submitButton.disabled = true;

        // Show loading back
        formContainer.style.display = 'none';
        loadingContainer.style.display = 'flex';
        updateProgress(10, null, 'Finalising...', 'Submitting policy confirmation');

        // Step 1: Confirm fields
        const confirmResponse = await fetch(`${BACKEND_URL}/api/email-fields`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({ filename: filename, email_fields: formData })
        });

        if (!confirmResponse.ok) throw new Error('Submission failed');
        updateProgress(40, null, 'Processing...', 'Generating analysis report');

        // Step 2: Process
        const processResponse = await fetch(`${BACKEND_URL}/api/process`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({ filename: filename })
        });

        if (!processResponse.ok) throw new Error('Processing failed');
        updateProgress(70, null, 'Publishing...', 'Uploading report to OneDrive');

        // Step 3: Poll for report
        let reportUrl = null;
        let attempts = 0;
        while (attempts < 20) {
            const pdfRes = await fetch(`${BACKEND_URL}/api/output-pdf`, {
                headers: { 'ngrok-skip-browser-warning': 'true' }
            });
            if (pdfRes.ok) {
                const pdfData = await pdfRes.json();
                reportUrl = pdfData.pdf_url;
                break;
            }
            await new Promise(r => setTimeout(r, 2000));
            attempts++;
        }

        if (reportUrl) {
            updateProgress(100, null, 'Complete!', 'Opening report...');
            window.open(reportUrl, '_blank');
        } else {
            updateProgress(100, null, 'Complete!', 'Report generated successfully');
        }

        setTimeout(() => location.reload(), 3000);

    } catch (error) {
        console.error(error);
        alert('Error: ' + error.message);
        loadingContainer.style.display = 'none';
        formContainer.style.display = 'block';
        submitButton.disabled = false;
    }
}
