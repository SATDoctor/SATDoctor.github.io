# SATDoctor.github.io

Marketing site and group-class booking calendar for SAT Doctor.

## Booking calendar

1. Deploy `apps-script/Code.gs` to your Google Sheet (see [apps-script/README.md](apps-script/README.md)).
2. Add group sessions on the **Slots** sheet (or use **SAT Doctor → Add sample group slots**).
3. Redeploy the web app and set `APPS_SCRIPT_URL` in `index.html` if the URL changed.

Students book from **Book a Group Class** on the homepage; confirmations are emailed with the Zoom link.