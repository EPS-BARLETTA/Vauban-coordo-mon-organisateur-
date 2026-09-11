export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const rawUrl = String(req.body?.url || "").trim();

    if (!rawUrl) {
      return res.status(400).json({ error: "Missing calendar URL" });
    }

    const url = new URL(rawUrl.replace(/^webcal:/i, "https:"));

    if (url.protocol !== "https:") {
      return res.status(400).json({ error: "HTTPS required" });
    }

    const host = url.hostname.toLowerCase();

    // Calendriers Microsoft / Outlook / iCloud autorisés
    const allowed =
      host === "outlook.office365.com" ||
      host === "outlook.office.com" ||
      host === "outlook.live.com" ||
      host.endsWith(".outlook.office365.com") ||
      host.endsWith(".office365.com") ||
      host.endsWith(".icloud.com");

    if (!allowed) {
      return res.status(403).json({ error: "Calendar host not allowed" });
    }

    const response = await fetch(url.toString(), {
      method: "GET",
      headers: {
        Accept: "text/calendar,text/plain;q=0.9,*/*;q=0.1",
        "User-Agent": "CoordoCalendar/1.0"
      },
      redirect: "follow"
    });

    if (!response.ok) {
      return res.status(502).json({
        error: "Calendar fetch failed",
        status: response.status
      });
    }

    const text = await response.text();

    if (text.length > 5_000_000) {
      return res.status(413).json({ error: "Calendar too large" });
    }

    // Déplie les lignes ICS coupées
    const unfolded = text.replace(/\r?\n[ \t]/g, "");

    const blocks = unfolded
      .split("BEGIN:VEVENT")
      .slice(1);

    const parseICSDate = value => {
      if (!value) {
        return { date: "", time: "" };
      }

      const clean = value.replace(/Z$/, "");

      // Journée entière : 20260911
      if (/^\d{8}$/.test(clean)) {
        return {
          date:
            `${clean.slice(0, 4)}-` +
            `${clean.slice(4, 6)}-` +
            `${clean.slice(6, 8)}`,
          time: ""
        };
      }

      // Date + heure : 20260911T173000
      if (/^\d{8}T\d{6}$/.test(clean)) {
        return {
          date:
            `${clean.slice(0, 4)}-` +
            `${clean.slice(4, 6)}-` +
            `${clean.slice(6, 8)}`,
          time:
            `${clean.slice(9, 11)}:` +
            `${clean.slice(11, 13)}`
        };
      }

      return { date: "", time: "" };
    };

    const events = blocks
      .map(block => {
        const body = block.split("END:VEVENT")[0] || "";
        const lines = body.split(/\r?\n/);

        const getValue = key => {
          const line = lines.find(
            l =>
              l.startsWith(key + ":") ||
              l.startsWith(key + ";")
          );

          if (!line) return "";

          return line
            .slice(line.indexOf(":") + 1)
            .trim();
        };

        const summary = getValue("SUMMARY")
          .replace(/\\n/g, " ")
          .replace(/\\,/g, ",")
          .replace(/\\;/g, ";");

        const location = getValue("LOCATION")
          .replace(/\\n/g, " ")
          .replace(/\\,/g, ",")
          .replace(/\\;/g, ";");

        const startLine =
          lines.find(l => l.startsWith("DTSTART")) || "";

        const endLine =
          lines.find(l => l.startsWith("DTEND")) || "";

        const startRaw = startLine.includes(":")
          ? startLine
              .slice(startLine.indexOf(":") + 1)
              .trim()
          : "";

        const endRaw = endLine.includes(":")
          ? endLine
              .slice(endLine.indexOf(":") + 1)
              .trim()
          : "";

        const start = parseICSDate(startRaw);
        const end = parseICSDate(endRaw);

        return {
          title: summary || "Événement",
          location,
          date: start.date,
          start: start.time,
          end: end.time
        };
      })
      .filter(event => event.date);

    res.setHeader(
      "Cache-Control",
      "private, max-age=60"
    );

    return res.status(200).json({
      events
    });

  } catch (error) {
    return res.status(400).json({
      error: "Invalid calendar request"
    });
  }
}
