// Outlook Email Extractor for Dev Console
(function() {
    let extractedEmails = [];
    let isMonitoring = false;
    
    // Debug function for DOM analysis
    function debugDOM() {
        console.log('🔍 DOM Analysis:');
        console.log('URL:', window.location.href);
        console.log('Title:', document.title);
        
        // Test all possible selectors
        const testSelectors = [
            '[data-testid="message-subject"]',
            '[role="heading"]',
            'h1', 'h2', 'h3',
            '.allowTextSelection',
            '[aria-label*="Von"]',
            '[aria-label*="From"]',
            '[title*="@"]',
            'time',
            '[datetime]',
            '.rps_813c',
            '[role="main"]'
        ];
        
        testSelectors.forEach(selector => {
            const elements = document.querySelectorAll(selector);
            if (elements.length > 0) {
                console.log(`✅ ${selector}: ${elements.length} element(s)`);
                elements.forEach((el, i) => {
                    if (i < 3) { // Only show first 3
                        console.log(`   [${i}]: "${el.textContent.trim().substring(0, 100)}..."`);
                    }
                });
            }
        });
    }
    
    // Extract email data
    function extractCurrentEmail() {
        const email = {};
        
        // Debug mode
        if (window.outlookDebug) {
            debugDOM();
        }
        
        // Extract subject - specific for Fielmann Outlook
        const subjectSelectors = [
            '.allowTextSelection:first-of-type', // The first .allowTextSelection is the subject
            '[data-testid="message-subject"]',
            'h1',
            'h2',
            'h3',
            '.rps_f409 span[role="heading"]',
            '.allowTextSelection[role="heading"]',
            'div[role="main"] [role="heading"]',
            '[aria-label*="Betreff"]',
            '[aria-label*="Subject"]'
        ];
        
        for (const selector of subjectSelectors) {
            const element = document.querySelector(selector);
            if (element && element.textContent.trim() && 
                !element.textContent.includes('Navigation pane') && 
                !element.textContent.includes('Inbox') &&
                !element.textContent.includes('Today')) {
                email.subject = element.textContent.trim();
                console.log(`📧 Subject found with: ${selector}`);
                break;
            }
        }
        
        // Extract sender - specific for Fielmann Outlook
        const senderSelectors = [
            '[aria-label*="From"]',
            '[aria-label*="Von"]',
            '[data-testid="message-header-from-button"] span',
            '.rps_f3cb .rps_f3cc',
            '[role="button"][title*="@"]'
        ];
        
        for (const selector of senderSelectors) {
            const element = document.querySelector(selector);
            if (element && element.textContent.trim() && 
                !element.textContent.includes('plano@fielmann.com')) {
                email.sender = element.textContent.trim();
                console.log(`👤 Sender found with: ${selector}`);
                break;
            }
        }
        
        // Extract date
        const dateSelectors = [
            '[data-testid="message-header-date-time"]',
            'time[datetime]',
            'time',
            '[aria-label*="Gesendet"]',
            '[aria-label*="Sent"]',
            '.rps_f3cb time'
        ];
        
        for (const selector of dateSelectors) {
            const element = document.querySelector(selector);
            if (element) {
                email.date = element.textContent.trim() || element.getAttribute('datetime');
                console.log(`📅 Date found with: ${selector}`);
                break;
            }
        }
        
        // Extract email content - specific for Fielmann Outlook
        const contentSelectors = [
            '.allowTextSelection:last-of-type', // The last .allowTextSelection is the content
            '[data-testid="message-body-content"]',
            '.rps_813c',
            '[role="main"] .allowTextSelection',
            'div[dir="ltr"]:not(:empty)',
            '.rps_813c .allowTextSelection',
            '[role="main"] div[class*="allowTextSelection"]'
        ];
        
        for (const selector of contentSelectors) {
            const element = document.querySelector(selector);
            if (element && element.textContent.trim().length > 50 && 
                !element.textContent.includes('Navigation pane') &&
                !element.textContent.includes('ReplyReply allForward')) {
                
                // Remove HTML comments and CSS
                let content = element.textContent.trim();
                content = content.replace(/<!--[\s\S]*?-->/g, '');
                content = content.replace(/\.[a-zA-Z0-9_]+ [a-zA-Z0-9_]*[\s\S]*?}/g, '');
                content = content.replace(/\s+/g, ' ');
                
                if (content.length > 50) {
                    email.content = content;
                    console.log(`📝 Content found with: ${selector}`);
                    break;
                }
            }
        }
        
        // Fallback: Collect all visible texts in main area
        if (!email.content) {
            const mainContent = document.querySelector('[role="main"]') || document.querySelector('main') || document.body;
            if (mainContent) {
                const allText = mainContent.textContent.trim();
                if (allText.length > 50) {
                    email.content = allText.substring(0, 2000);
                    console.log('📝 Fallback content used');
                }
            }
        }
        
        return email;
    }
    
    // Check if an email is open
    function isEmailOpen() {
        const hasMainContent = document.querySelector('[role="main"]');
        const hasEmailContent = document.querySelector('.allowTextSelection');
        
        // Extended URL recognition for different folders
        const isInEmailView = window.location.href.includes('/id/') || 
                             window.location.href.includes('/inbox/') ||
                             window.location.href.includes('/folders/');
        
        if (window.outlookDebug) {
            console.log('🔍 Email detection:');
            console.log('Main Content:', !!hasMainContent);
            console.log('Email Content:', !!hasEmailContent);
            console.log('Email View:', isInEmailView);
            console.log('URL:', window.location.href);
        }
        
        return hasMainContent && hasEmailContent && isInEmailView;
    }
    
    // Save email
    function saveEmail() {
        if (!isEmailOpen()) {
            console.log('❌ No email open');
            return;
        }
        
        const email = extractCurrentEmail();
        
        if (email.subject && email.content) {
            email.timestamp = new Date().toISOString();
            email.url = window.location.href;
            
            // Avoid duplicates
            const isDuplicate = extractedEmails.some(existing => 
                existing.subject === email.subject && 
                existing.sender === email.sender &&
                existing.url === email.url
            );
            
            if (!isDuplicate) {
                extractedEmails.push(email);
                console.log('📧 Email saved:', email.subject);
                console.log('📊 Total emails:', extractedEmails.length);
            } else {
                console.log('⚠️ Duplicate already exists');
            }
        } else {
            console.log('❌ Incomplete email data:', email);
        }
    }
    
    // Start monitoring
    function startMonitoring() {
        if (isMonitoring) return;
        
        isMonitoring = true;
        console.log('🔍 Email monitoring started');
        
        // Capture current email
        setTimeout(saveEmail, 1000);
        
        // Observer for DOM changes
        const observer = new MutationObserver(function(mutations) {
            let shouldCheck = false;
            mutations.forEach(function(mutation) {
                if (mutation.addedNodes.length > 0) {
                    shouldCheck = true;
                }
            });
            
            if (shouldCheck) {
                setTimeout(saveEmail, 500);
            }
        });
        
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
        
        // Monitor click events
        document.addEventListener('click', function() {
            setTimeout(saveEmail, 1000);
        });
        
        window.outlookEmailExtractor = {
            observer: observer,
            stop: function() {
                observer.disconnect();
                isMonitoring = false;
                console.log('⏹️ Monitoring stopped');
            }
        };
    }
    
    // Export functions
    window.exportEmails = {
        // Download as JSON
        downloadJSON: function() {
            const dataStr = JSON.stringify(extractedEmails, null, 2);
            const dataBlob = new Blob([dataStr], {type: 'application/json'});
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'outlook_emails_' + new Date().toISOString().split('T')[0] + '.json';
            link.click();
            URL.revokeObjectURL(url);
        },
        
        // Download as CSV
        downloadCSV: function() {
            if (extractedEmails.length === 0) {
                console.log('No emails to export');
                return;
            }
            
            const headers = ['Timestamp', 'Subject', 'Sender', 'Date', 'Content'];
            const csvContent = [
                headers.join(','),
                ...extractedEmails.map(email => [
                    `"${email.timestamp || ''}"`,
                    `"${(email.subject || '').replace(/"/g, '""')}"`,
                    `"${(email.sender || '').replace(/"/g, '""')}"`,
                    `"${(email.date || '').replace(/"/g, '""')}"`,
                    `"${(email.content || '').replace(/"/g, '""').substring(0, 500)}"`
                ].join(','))
            ].join('\n');
            
            const dataBlob = new Blob([csvContent], {type: 'text/csv'});
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'outlook_emails_' + new Date().toISOString().split('T')[0] + '.csv';
            link.click();
            URL.revokeObjectURL(url);
        },
        
        // Show emails
        show: function() {
            console.table(extractedEmails);
            return extractedEmails;
        },
        
        // Clear emails
        clear: function() {
            extractedEmails = [];
            console.log('🗑️ All saved emails cleared');
        },
        
        // Enable debug mode
        debug: function() {
            window.outlookDebug = true;
            console.log('🐛 Debug mode enabled');
            debugDOM();
        },
        
        // Manually capture current email
        capture: function() {
            console.log('🔍 Attempting to capture current email...');
            
            if (!isEmailOpen()) {
                console.log('❌ No email open. Please open an email first.');
                return null;
            }
            
            const email = extractCurrentEmail();
            
            if (email.subject && email.content) {
                email.timestamp = new Date().toISOString();
                email.url = window.location.href;
                
                const isDuplicate = extractedEmails.some(existing => 
                    existing.subject === email.subject && 
                    existing.sender === email.sender &&
                    existing.url === email.url
                );
                
                if (!isDuplicate) {
                    extractedEmails.push(email);
                    console.log('✅ Email captured:', email);
                    console.log('📊 Total emails:', extractedEmails.length);
                } else {
                    console.log('⚠️ Duplicate detected, not saved');
                }
            } else {
                console.log('❌ Incomplete email data found:');
                console.log('Subject:', email.subject || 'MISSING');
                console.log('Sender:', email.sender || 'MISSING');
                console.log('Content:', email.content ? 'AVAILABLE' : 'MISSING');
            }
            
            return email;
        }
    };
    
    // Auto-start
    startMonitoring();
    
    // Show help
    console.log(`
🚀 Outlook Email Extractor started!

Available commands:
- exportEmails.show()         - Show all saved emails
- exportEmails.capture()      - Manually capture current email
- exportEmails.debug()        - Enable debug mode
- exportEmails.downloadJSON() - Download as JSON file
- exportEmails.downloadCSV()  - Download as CSV file
- exportEmails.clear()        - Clear all saved emails
- outlookEmailExtractor.stop() - Stop monitoring

The script automatically captures emails when you open them.
If it doesn't work, use exportEmails.debug() and exportEmails.capture()
Saved emails: ${extractedEmails.length}
    `);
    
})();
