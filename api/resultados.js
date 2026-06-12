// Función serverless de Vercel — corre en el servidor, no en el navegador.
// Oculta la API key de football-data.org y devuelve solo los datos que necesitamos.
//
// Configuración necesaria en Vercel:
//   Settings → Environment Variables → FOOTBALL_DATA_TOKEN = <tu API key>

export default async function handler(req, res) {
  const token = process.env.FOOTBALL_DATA_TOKEN;

  if (!token) {
    return res.status(500).json({ error: "Falta configurar FOOTBALL_DATA_TOKEN en Vercel" });
  }

  try {
    const r = await fetch("https://api.football-data.org/v4/competitions/WC/matches", {
      headers: { "X-Auth-Token": token },
    });

    if (!r.ok) {
      const txt = await r.text();
      return res.status(r.status).json({ error: `football-data.org respondió ${r.status}`, detail: txt });
    }

    const data = await r.json();
    const matches = data.matches || [];

    // Solo partidos que ya tienen resultado de tiempo regular
    const results = matches
      .filter(m => m.score?.fullTime?.home !== null && m.score?.fullTime?.home !== undefined
                 && m.score?.fullTime?.away !== null && m.score?.fullTime?.away !== undefined)
      .map(m => ({
        local: m.homeTeam?.name || "",
        visitante: m.awayTeam?.name || "",
        gl: m.score.fullTime.home,
        gv: m.score.fullTime.away,
        status: m.status,
        utcDate: m.utcDate,
      }));

    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Cache-Control", "s-maxage=15, stale-while-revalidate=30");
    return res.status(200).json({ results, total: matches.length });

  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
