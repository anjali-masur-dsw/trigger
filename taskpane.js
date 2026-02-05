let mailboxItem = null;
let filename = '';
let processingType = ''; // 'claims' or 'underwriting'

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        mailboxItem = Office.context.mailbox.item;
        console.log('Office.js initialized successfully');

        // Automatically trigger the flow when taskpane opens
        triggerFlowAndLoadForm();
    }
});

async function triggerFlowAndLoadForm() {
    const loadingContainer = document.getElementById('loadingContainer');
    const formContainer = document.getElementById('formContainer');
    const loadingText = document.querySelector('.loading-text');
    const loadingSubtext = document.querySelector('.loading-subtext');

    try {
        // Show progress notification in email view
        Office.context.mailbox.item.notificationMessages.addAsync(
            "progress",
            {
                type: "progressIndicator",
                message: "Analyzing email and sending to workflow..."
            }
        );

        // Step 1: Get email data
        loadingText.textContent = 'Collecting email data...';
        const emailData = await getEmailData();
        console.log('Email data collected:', emailData);

        // Step 2: Trigger Power Automate flow
        loadingText.textContent = 'Triggering Power Automate flow...';
        loadingSubtext.textContent = 'Sending email data for processing';

        const flowUrl = "https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/89c12382226642a4907cd110e9e7ab87/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Nbz7sUIbNoHlSBt_KVnF3CFKCCf9lPYn-LbIxZsWouA";

        const response = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(emailData)
        });

        Office.context.mailbox.item.notificationMessages.removeAsync("progress");

        if (!response.ok) {
            const err = await response.text();
            throw new Error(`HTTP ${response.status}: ${err}`);
        }

        console.log('Flow triggered successfully');

        // Step 3: Poll for extracted form data
        // Show notification for form data extraction
        Office.context.mailbox.item.notificationMessages.addAsync(
            "formProcessing",
            {
                type: "informationalMessage",
                message: "Extracting form data from attachments...",
                icon: "Icon.80x80",
                persistent: true
            }
        );

        loadingText.textContent = 'Extracting form data from attachments...';
        loadingSubtext.textContent = 'This may take a few moments';

        let extractedData = null;
        let pollingAttempts = 0;
        const maxPollingAttempts = 60; // 5 minutes max

        while (pollingAttempts < maxPollingAttempts) {
            try {
                // Note: At this stage we don't know the processing type yet,
                // so we try /api/pending first (underwriting default)
                // The processing_type will be determined once we get the data
                const pendingResponse = await fetch("https://corinne-unstudded-uneugenically.ngrok-free.dev/api/pending", {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });

                if (pendingResponse.ok) {
                    const data = await pendingResponse.json();
                    console.log("Pending API response:", data);

                    // Check for unified API response format (single object)
                    if (data.success && data.session_id && data.extracted_data) {
                        extractedData = data.extracted_data;
                        // Add session_id to the data for tracking
                        extractedData.session_id = data.session_id;
                        extractedData.processing_type = data.processing_type;

                        console.log("Extracted data received:", extractedData);
                        break;
                    }
                    // Fallback for legacy format (files array)
                    else if (data.success && data.files && data.files.length > 0) {
                        extractedData = data.files[0];
                        console.log("Extracted data received (legacy):", extractedData);
                        break;
                    }
                }
            } catch (pollError) {
                console.warn("Polling attempt failed:", pollError.message);
            }

            // Wait 5 seconds before next attempt
            await new Promise(resolve => setTimeout(resolve, 5000));
            pollingAttempts++;

            if (pollingAttempts % 6 === 0) {
                loadingSubtext.textContent = `Still processing... (${Math.floor(pollingAttempts / 12)} minute${Math.floor(pollingAttempts / 12) > 1 ? 's' : ''})`;
            }
        }

        if (!extractedData) {
            Office.context.mailbox.item.notificationMessages.removeAsync("formProcessing");
            throw new Error('Timeout: Form data extraction did not complete in time');
        }

        // Remove processing notification
        Office.context.mailbox.item.notificationMessages.removeAsync("formProcessing");

        // Show success notification
        Office.context.mailbox.item.notificationMessages.addAsync(
            "formSuccess",
            {
                type: "informationalMessage",
                message: "Form data extracted! Opening form in taskpane...",
                icon: "Icon.80x80",
                persistent: false
            }
        );

        // Step 4: Populate form with extracted data
        loadingText.textContent = 'Loading form...';
        loadingSubtext.textContent = 'Preparing your insurance policy information';

        populateForm(extractedData);

        // Show form, hide loading
        loadingContainer.style.display = 'none';
        formContainer.style.display = 'block';

    } catch (error) {
        console.error('Error:', error);

        // Remove any pending notifications
        Office.context.mailbox.item.notificationMessages.removeAsync("progress");
        Office.context.mailbox.item.notificationMessages.removeAsync("formProcessing");

        // Show error notification in email view
        Office.context.mailbox.item.notificationMessages.addAsync(
            "error",
            {
                type: "errorMessage",
                message: "Processing failed: " + error.message
            }
        );

        loadingText.textContent = 'Error occurred';
        loadingSubtext.innerHTML = `<div class="error-message">${error.message}</div>`;
    }
}

function populateForm(extractedData) {
    // Store filename for later submission
    filename = extractedData.filename || '';

    // Detect processing type based on filename prefix
    if (filename) {
        if (filename.toLowerCase().startsWith('c')) {
            processingType = 'claims';
            console.log('Processing type detected: CLAIMS');
        } else if (filename.toLowerCase().startsWith('acord')) {
            processingType = 'underwriting';
            console.log('Processing type detected: UNDERWRITING');
        } else {
            processingType = 'underwriting'; // default
            console.log('Processing type defaulted to: UNDERWRITING');
        }
    }

    // Get email_fields from the extracted data
    const data = extractedData.email_fields || extractedData.extracted_data || {};

    console.log("Populating form with data:", data);

    // Populate form fields
    if (data.broker_email) document.getElementById('brokerEmail').value = data.broker_email;
    if (data.broker_name) document.getElementById('brokerName').value = data.broker_name;
    if (data.underwriter_email) document.getElementById('underwriterEmail').value = data.underwriter_email;
    if (data.underwriter_name) document.getElementById('underwriterName').value = data.underwriter_name;
    if (data.policy_number) document.getElementById('policyNumber').value = data.policy_number;
    if (data.broker_agency_name) document.getElementById('agencyName').value = data.broker_agency_name;
    if (data.broker_agency_id || data.agency_id) document.getElementById('agencyId').value = data.broker_agency_id || data.agency_id;
    if (data.email_summary) document.getElementById('emailSummary').value = data.email_summary;
    if (data.comments) document.getElementById('comments').value = data.comments;

    // Set timestamp
    const timestampField = document.getElementById('timestamp');
    if (extractedData.detected_at) {
        // Format detected_at timestamp (remove seconds)
        const detectedDate = new Date(extractedData.detected_at);
        timestampField.value = detectedDate.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    } else {
        const now = new Date();
        timestampField.value = now.toLocaleString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    }
}

// Collapse form permanently after successful submission
function collapseForm() {
    const formContainer = document.getElementById('formContainer');
    const formBody = document.querySelector('.form-body');
    const submitSection = document.querySelector('.submit-section');
    const header = document.querySelector('.header');
    const successMessage = document.getElementById('successMessage');

    console.log('Collapsing form...');

    // Hide form body, submit section, and success message
    if (formBody) {
        formBody.style.display = 'none';
        console.log('Form body hidden');
    }
    if (submitSection) {
        submitSection.style.display = 'none';
        console.log('Submit section hidden');
    }
    if (successMessage) {
        successMessage.classList.remove('show');
        successMessage.style.display = 'none';
        console.log('Success message hidden');
    }

    // Add collapsed state styling
    if (formContainer) {
        formContainer.style.transition = 'all 0.3s ease';
    }

    // Create or update success summary
    let successSummary = document.getElementById('successSummary');
    if (!successSummary) {
        successSummary = document.createElement('div');
        successSummary.id = 'successSummary';
        successSummary.className = 'success-summary';
        successSummary.innerHTML = `
            <div class="success-summary-content">
                <div class="success-icon">✓</div>
                <div class="success-text">
                    <h3>Form Submitted Successfully</h3>
                    <p>Your insurance policy information has been processed and the report has been generated.</p>
                </div>
            </div>
        `;

        // Insert after header
        if (header && header.parentNode) {
            header.parentNode.insertBefore(successSummary, header.nextSibling);
        }
    }

    successSummary.style.display = 'block';
    console.log('Success summary displayed');
}

async function getEmailData() {
    const item = mailboxItem;
    if (!item) throw new Error("No email context");

    const subject = await getSubject(item);
    const body = await getBody(item);
    const attachments = await getAttachmentContents(item);

    const itemId = item.itemId || "";
    const conversationId = item.conversationId || "";

    console.log("Item ID:", itemId);
    console.log("Conversation ID:", conversationId);

    return {
        triggeredAt: new Date().toISOString(),
        userEmail: Office.context.mailbox.userProfile.emailAddress || "",
        subject: subject,
        body: body,
        from: getSender(item),
        receivedDateTime: item.dateTimeCreated || item.dateTimeModified || new Date().toISOString(),
        internetMessageId: item.internetMessageId || "",
        itemId: itemId,
        conversationId: conversationId,
        hasAttachments: attachments.length > 0,
        attachmentCount: attachments.length,
        attachments: attachments
    };
}

function getSubject(item) {
    return new Promise(resolve => {
        if (typeof item.subject === "string") {
            resolve(item.subject);
        } else {
            item.subject.getAsync(result =>
                resolve(result.status === Office.AsyncResultStatus.Succeeded
                    ? result.value
                    : "Subject unavailable")
            );
        }
    });
}

function getBody(item) {
    return new Promise(resolve => {
        item.body.getAsync(Office.CoercionType.Text, result =>
            resolve(result.status === Office.AsyncResultStatus.Succeeded
                ? result.value
                : "Body unavailable")
        );
    });
}

function getSender(item) {
    if (!item.from) return "Unknown";
    if (typeof item.from === "string") return item.from;
    if (item.from.emailAddress) return item.from.emailAddress;
    if (item.from.displayName) return item.from.displayName;
    return "Unknown";
}

async function getAttachmentContents(item) {
    if (!item.attachments || item.attachments.length === 0) {
        return [];
    }

    const results = [];

    for (const att of item.attachments) {
        console.log(`Processing attachment: ${att.name}`);

        try {
            await new Promise(resolve => {
                item.getAttachmentContentAsync(att.id, res => {
                    if (res.status === Office.AsyncResultStatus.Succeeded) {
                        results.push({
                            id: att.id,
                            name: att.name,
                            contentType: att.contentType,
                            size: att.size,
                            content: res.value.content,
                            format: res.value.format
                        });
                    } else {
                        results.push({
                            id: att.id,
                            name: att.name,
                            error: res.error ? res.error.message : 'Unknown error'
                        });
                    }
                    resolve();
                });
            });
        } catch (error) {
            console.error(`Exception retrieving ${att.name}:`, error);
            results.push({
                id: att.id,
                name: att.name,
                error: error.message
            });
        }
    }

    return results;
}

// Generate PDF from form data
async function generateFormPDF(formData) {
    return new Promise((resolve) => {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        // Add header
        doc.setFillColor(249, 168, 37);
        doc.rect(0, 0, 210, 40, 'F');

        doc.setTextColor(255, 255, 255);
        doc.setFontSize(22);
        doc.setFont(undefined, 'bold');
        doc.text('Insurance Policy Information', 105, 20, { align: 'center' });

        doc.setFontSize(12);
        doc.setFont(undefined, 'normal');
        doc.text('Submitted Form Response', 105, 30, { align: 'center' });

        // Reset text color for body
        doc.setTextColor(0, 0, 0);
        let yPos = 50;

        // Add form fields
        const fields = [
            { label: "Sender's Email", value: formData.broker_email },
            { label: "Sender's Name", value: formData.broker_name },
            { label: "Receiver's Email", value: formData.underwriter_email },
            { label: "Receiver's Name", value: formData.underwriter_name },
            { label: "Policy Number", value: formData.policy_number },
            { label: "Agency Name", value: formData.broker_agency_name },
            { label: "Agency/Broker ID", value: formData.broker_agency_id },
            { label: "Email Summary", value: formData.email_summary, multiline: true },
            { label: "Comments", value: formData.comments, multiline: true },
            { label: "Timestamp", value: formData.timestamp }
        ];

        fields.forEach(field => {
            if (field.value) {
                doc.setFontSize(11);
                doc.setFont(undefined, 'bold');
                doc.text(field.label + ':', 20, yPos);

                doc.setFont(undefined, 'normal');
                doc.setFontSize(10);

                if (field.multiline && field.value.length > 60) {
                    // Handle long text with wrapping
                    const lines = doc.splitTextToSize(field.value, 170);
                    yPos += 6;
                    lines.forEach(line => {
                        if (yPos > 270) {
                            doc.addPage();
                            yPos = 20;
                        }
                        doc.text(line, 20, yPos);
                        yPos += 5;
                    });
                    yPos += 3;
                } else {
                    yPos += 6;
                    doc.text(field.value, 20, yPos);
                    yPos += 8;
                }

                if (yPos > 270) {
                    doc.addPage();
                    yPos = 20;
                }
            }
        });

        // Add footer
        const pageCount = doc.internal.getNumberOfPages();
        for (let i = 1; i <= pageCount; i++) {
            doc.setPage(i);
            doc.setFontSize(9);
            doc.setTextColor(128, 128, 128);
            doc.text(`Page ${i} of ${pageCount}`, 105, 285, { align: 'center' });

        }

        // Convert to base64
        const pdfBase64 = doc.output('dataurlstring').split(',')[1];
        resolve(pdfBase64);
    });
}

// Handle form submission
document.addEventListener('DOMContentLoaded', function () {
    const insuranceForm = document.getElementById('insuranceForm');
    if (insuranceForm) {
        insuranceForm.addEventListener('submit', handleFormSubmit);
    }
});

async function handleFormSubmit(e) {
    e.preventDefault();

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

    const submitButton = document.querySelector('.submit-button');
    const successMessage = document.getElementById('successMessage');

    try {
        submitButton.disabled = true;
        submitButton.textContent = 'Generating PDF...';

        // Generate PDF from form data
        const pdfBase64 = await generateFormPDF(formData);
        console.log('PDF generated successfully');

        submitButton.textContent = 'Submitting...';

        // Step 3: Confirm email fields with PDF
        console.log('Confirming email fields...');
        const apiPrefix = processingType === 'claims' ? '/claims-api' : '/api';
        console.log(`Using API prefix: ${apiPrefix}`);
        
        const confirmResponse = await fetch(`https://corinne-unstudded-uneugenically.ngrok-free.dev${apiPrefix}/email-fields`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                filename: filename,
                email_fields: formData,
                form_pdf: pdfBase64  // Include the PDF in base64 format
            })
        });

        if (!confirmResponse.ok) {
            throw new Error(`Email fields confirmation failed: ${confirmResponse.status}`);
        }

        console.log('Email fields confirmed');

        // Step 4: Process the file
        console.log('Processing file...');
        submitButton.textContent = 'Processing...';

        const processResponse = await fetch(`https://corinne-unstudded-uneugenically.ngrok-free.dev${apiPrefix}/process`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                filename: filename
            })
        });

        if (!processResponse.ok) {
            throw new Error(`Processing failed: ${processResponse.status}`);
        }

        const result = await processResponse.json();
        console.log('Processing successful:', result);

        successMessage.textContent = '✓ Form submitted successfully! Generating report...';
        successMessage.classList.add('show');
        submitButton.textContent = 'Generating Report...';

        // Step 5: Poll for PDF report - start immediately with shorter interval
        console.log('Waiting for PDF report...');
        let pdfReady = false;
        let pdfPollingAttempts = 0;
        const maxPdfPollingAttempts = 60; // 2 minutes max (60 attempts * 2 seconds)

        while (pdfPollingAttempts < maxPdfPollingAttempts && !pdfReady) {
            try {
                const pdfResponse = await fetch(`https://corinne-unstudded-uneugenically.ngrok-free.dev${apiPrefix}/output-pdf`, {
                    headers: {
                        'ngrok-skip-browser-warning': 'true',
                        'Accept': 'application/json'
                    }
                });

                if (pdfResponse.status === 200) {
                    const pdfData = await pdfResponse.json();
                    console.log('PDF is ready!', pdfData);

                    const reportUrl = pdfData.pdf_url;

                    if (reportUrl) {
                        console.log('Opening report:', reportUrl);

                        // Open report in new window
                        try {
                            const newWindow = window.open(reportUrl, '_blank', 'noopener,noreferrer');

                            if (newWindow && !newWindow.closed && typeof newWindow.closed !== 'undefined') {
                                console.log('Report opened successfully in new window');
                                successMessage.textContent = '✓ Email processed successfully! Report opened.';
                            } else {
                                console.log('Window blocked, showing link');
                                successMessage.innerHTML = `✓ Email processed successfully! <a href="${reportUrl}" target="_blank" style="color: #0078d4; text-decoration: underline; font-weight: bold;">Click here to open report</a>`;
                            }
                        } catch (openError) {
                            console.error('Error opening report:', openError);
                            successMessage.innerHTML = `✓ Email processed successfully! <a href="${reportUrl}" target="_blank" style="color: #0078d4; text-decoration: underline; font-weight: bold;">Click here to open report</a>`;
                        }
                    } else {
                        successMessage.textContent = '✓ Email processed successfully! Report URL not available.';
                    }

                    pdfReady = true;
                    break;
                }
            } catch (pollError) {
                console.warn('PDF polling attempt failed:', pollError.message);
            }

            await new Promise(resolve => setTimeout(resolve, 2000));
            pdfPollingAttempts++;
        }

        if (!pdfReady) {
            successMessage.textContent = '✓ Email processed successfully! Report is still processing...';
        }

        // Show persistent notification in email view
        Office.context.mailbox.item.notificationMessages.addAsync(
            "processSuccess",
            {
                type: "informationalMessage",
                message: "Email processed successfully! ✓",
                icon: "Icon.80x80",
                persistent: true
            }
        );

        // Keep button as "Processed" and disabled permanently
        submitButton.textContent = 'Processed';
        submitButton.disabled = true;

        // Close the taskpane after showing success notification
        setTimeout(() => {
            console.log('Closing taskpane...');
            Office.context.ui.closeContainer();
        }, 2000);

    } catch (error) {
        console.error('Error submitting form:', error);

        successMessage.textContent = '✗ Submission failed: ' + error.message;
        successMessage.style.background = '#fde7e9';
        successMessage.style.color = '#a80000';
        successMessage.style.borderLeft = '4px solid #a80000';
        successMessage.classList.add('show');

        submitButton.disabled = false;
        submitButton.textContent = 'Submit';

        setTimeout(() => {
            successMessage.classList.remove('show');
            successMessage.style.background = '#dff6dd';
            successMessage.style.color = '#0b7815';
            successMessage.style.borderLeft = '4px solid #0b7815';
        }, 5000);
    }
}
