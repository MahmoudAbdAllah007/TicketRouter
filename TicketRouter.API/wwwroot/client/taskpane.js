(function () {
    const statusEl = () => document.getElementById('status');

    Office.onReady(() => {
        const item = Office.context.mailbox.item;
        const subject = item && item.subject ? item.subject : '';
        const m = (subject || '').match(/TrackingID#(\d{15,})/);
        if (m) document.getElementById('txtTicket').value = m[1];

        document.getElementById('btnRoute').addEventListener('click', routeSelected);
        document.getElementById('btnClose').addEventListener('click', closeTicket);
    });

    async function getApiToken() {
        return new Promise((resolve, reject) => {
            Office.auth.getAccessToken({ allowSignInPrompt: true }, res => {
                if (res.status === 'succeeded') resolve(res.value);
                else reject(res.error);
            });
        });
    }

    function getItemId() {
        const item = Office.context.mailbox.item;
        return item.itemId || item.restId;
    }

    function getSubject() {
        const item = Office.context.mailbox.item;
        return item.subject || '';
    }

    async function routeSelected() {
        try {
            const subject = getSubject();
            const ticketId = (subject.match(/TrackingID#(\d{15,})/) || [])[1] || document.getElementById('txtTicket').value.trim();
            if (!ticketId) return setStatus('No TrackingID found.', true);

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
            const data = await r.json();
            if (!r.ok) throw new Error(data.error || 'Routing failed.');
            setStatus(`Routed to folder for ${data.ticketId}`);
        } catch (e) { setStatus(e.message, true); }
    }

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

    function setStatus(msg, err) {
        statusEl().textContent = msg;
        statusEl().style.color = err ? '#b91c1c' : '#16a34a';
    }
})();