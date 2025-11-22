(function () {
    const statusEl = () => document.getElementById('status');

    // Initialize the Office.js library
    Office.onReady(() => {
        //get the current item's subject
        const item = Office.context.mailbox.item;
        const subject = item && item.subject ? item.subject : '';
        //auto-fill ticket ID if found in subject
        const m = (subject || '').match(/TrackingID#(\d{15,})/);
        if (m) document.getElementById('txtTicket').value = m[1];
        //set up button click handlers
        document.getElementById('btnRoute').addEventListener('click', routeSelected);
        document.getElementById('btnClose').addEventListener('click', closeTicket);
    });

    //get an API token using Office.js SSO
    async function getApiToken() {
        //wrap Office.auth.getAccessToken in a Promise
        return new Promise((resolve, reject) => {
            ///call Office.js to get the SSO token
            Office.auth.getAccessToken({ allowSignInPrompt: true }, res => {
                if (res.status === 'succeeded') resolve(res.value);
                else reject(res.error);
            });
        });
    }
    //get the current item's ID — returns the current message’s ID (REST ID fallback)
    function getItemId() {
        const item = Office.context.mailbox.item;
        return item.itemId || item.restId;
    }

    //get the current item's subject
    function getSubject() {
        const item = Office.context.mailbox.item;
        return item.subject || '';
    }

    //route the selected email to the appropriate folder based on ticket ID
    async function routeSelected() {
        try {
            const subject = getSubject();
            const ticketId = (subject.match(/TrackingID#(\d{15,})/) || [])[1] || document.getElementById('txtTicket').value.trim();
            if (!ticketId) return setStatus('No TrackingID found.', true);

            //call the API to route the email
            const token = await getApiToken();
            const r = await fetch('/api/tickets/routeSelected', {
                method: 'POST',
                headers: { 'content-type': 'application/json', 'authorization': `Bearer ${token}` },
                body: JSON.stringify({
                    messageId: getItemId(),
                    subject,
                    shortName: document.getElementById('txtShort').value
                })
            });
            //parse the response
            const data = await r.json();
            if (!r.ok) throw new Error(data.error || 'Routing failed.');
            setStatus(`Routed to folder for ${data.ticketId}`);
        } catch (e) { setStatus(e.message, true); }
    }

    //close the ticket with the given ticket ID
    async function closeTicket() {
        try {
            const ticketId = document.getElementById('txtTicket').value.trim();
            if (!ticketId) return setStatus('Enter the TrackingID to close.', true);

            const token = await getApiToken();
            const r = await fetch('/api/tickets/close', {
                method: 'POST',
                headers: { 'content-type': 'application/json', 'authorization': `Bearer ${token}` },
                body: JSON.stringify({ ticketId })
            });
            const data = await r.json();
            if (!r.ok) throw new Error(data.error || 'Close failed.');
            setStatus(data.status);
        } catch (e) { setStatus(e.message, true); }
    }

    //update status message — updates a status element with success/error coloring
    function setStatus(msg, err) {
        statusEl().textContent = msg;
        statusEl().style.color = err ? '#b91c1c' : '#16a34a';
    }
})();