const Imap = require('node-imap');
const simpleParser = require('mailparser').simpleParser;

async function get_access_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();
    try {
        const data = JSON.parse(responseText);
        return data.access_token;
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

async function graph_api(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'scope': 'https://graph.microsoft.com/.default'
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();
    try {
        const data = JSON.parse(responseText);
        if (data.scope && data.scope.indexOf && data.scope.indexOf('https://graph.microsoft.com/Mail.ReadWrite') !== -1) {
            return { access_token: data.access_token, status: true };
        }
        return { access_token: data.access_token, status: false };
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

async function get_emails_graph(access_token, mailbox, limit) {
    if (!access_token) return [];
    try {
        // Use $top and order by receivedDateTime desc to get latest emails
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=${limit}&$orderby=receivedDateTime desc`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Authorization': `Bearer ${access_token}`
            }
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Graph API error: ${response.status} ${errorText}`);
        }

        const data = await response.json();
        const emails = data.value || [];
        return emails.map(item => ({
            send: item.from && item.from.emailAddress ? item.from.emailAddress.address : '',
            subject: item.subject,
            text: item.bodyPreview,
            html: item.body && item.body.content,
            date: item.receivedDateTime || item.createdDateTime
        }));
    } catch (err) {
        console.error('Error fetching from graph:', err);
        return [];
    }
}

module.exports = async (req, res) => {
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;
    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({ error: 'Authentication failed. Please provide valid credentials or contact administrator for access.' });
    }

    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, limit = 10 } = params;
    limit = parseInt(limit, 10) || 10;

    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    try {
        const graphResult = await graph_api(refresh_token, client_id);
        if (graphResult.status) {
            if (mailbox !== 'INBOX' && mailbox !== 'Junk') mailbox = 'inbox';
            if (mailbox === 'INBOX') mailbox = 'inbox';
            if (mailbox === 'Junk') mailbox = 'junkemail';

            const emails = await get_emails_graph(graphResult.access_token, mailbox, limit);
            return res.status(200).json(emails);
        }

        // Fallback to IMAP
        const access_token = await get_access_token(refresh_token, client_id);
        const authString = generateAuthString(email, access_token);

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: { rejectUnauthorized: false }
        });

        const emailList = [];

        imap.once('ready', async () => {
            try {
                await new Promise((resolve, reject) => imap.openBox(mailbox, true, (err, box) => err ? reject(err) : resolve(box)));

                const results = await new Promise((resolve, reject) => imap.search(['ALL'], (err, results) => err ? reject(err) : resolve(results)));

                if (!results || results.length === 0) {
                    imap.end();
                    return res.status(200).json([]);
                }

                // pick the latest `limit` messages
                const idsToFetch = results.slice(-limit);

                const f = imap.fetch(idsToFetch, { bodies: '' });

                f.on('message', (msg, seqno) => {
                    msg.on('body', (stream) => {
                        simpleParser(stream, (err, mail) => {
                            if (err) {
                                console.error('parse mail err', err);
                                return;
                            }
                            emailList.push({
                                send: mail.from ? mail.from.text : '',
                                subject: mail.subject,
                                text: mail.text,
                                html: mail.html,
                                date: mail.date
                            });
                        });
                    });
                });

                f.once('error', (err) => {
                    console.error('Fetch error:', err);
                    imap.end();
                    return res.status(500).json({ error: 'Fetch error', details: err.message });
                });

                f.once('end', () => {
                    // results are in ascending order by sequence; ensure we return latest-first
                    emailList.sort((a, b) => new Date(b.date) - new Date(a.date));
                    imap.end();
                    return res.status(200).json(emailList);
                });
            } catch (err) {
                console.error('IMAP ready error:', err);
                imap.end();
                return res.status(500).json({ error: err.message });
            }
        });

        imap.once('error', (err) => {
            console.error('IMAP connection error:', err);
            return res.status(500).json({ error: err.message });
        });

        imap.connect();

    } catch (error) {
        console.error('Error:', error);
        return res.status(500).json({ error: error.message });
    }
};
