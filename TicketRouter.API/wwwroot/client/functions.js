(function () {
    Office.onReady(() => { });

    async function getApiToken() {
        return new Promise((resolve, reject) => {
            Office.auth.getAccessToken({ allowSignInPrompt: true }, res => {
                if (res.status === 'succeeded') resolve(res.value);
                else reject(res.error);
            });
        });
    }

    // Ribbon button: Close Ticket
    window.closeTicket = async function (event) {
        try {
            const subject = (Office.context.mailbox.item && Office.context.mailbox.item.subject) || '';
            const m = subject.match(/TrackingID#(\d{15,})/);
            if (!m) return event.completed();

            const token = await getApiToken();
            await fetch('/api/tickets/close', {
                method: 'POST',
                headers: { 'content-type': 'application/json', 'authorization': `Bearer ${token}` },
                body: JSON.stringify({ ticketId: m[1] })
            });
        } catch (e) { /* swallow */ }
        finally { event.completed(); }
    }

    // Smart Alerts: OnMessageSend
    window.onMessageSend = async function (event) {
        try {
            const compose = Office.context.mailbox.item;
            const subject = compose.subject || '';
            const m = subject.match(/TrackingID#(\d{15,})/);
            if (!m) return event.completed({ allowEvent: true });

            await new Promise(resolve => compose.saveAsync(() => resolve()));
            const token = await getApiToken();
            await fetch('/api/tickets/sentRoute', {
                method: 'POST',
                headers: { 'content-type': 'application/json', 'authorization': `Bearer ${token}` },
                body: JSON.stringify({ messageId: compose.itemId, subject })
            });
            return event.completed({ allowEvent: true });
        } catch (e) {
            return event.completed({ allowEvent: true });
        }
    }
})();