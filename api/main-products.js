export default async function handler(req, res) {
  try {
    const upstream = await fetch('https://www.fairplay142.com/', {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; ADpicker/1.0)',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8',
      },
    });
    if (!upstream.ok) {
      return res.status(upstream.status).json({ error: `사이트 응답 오류 (${upstream.status})` });
    }
    const html = await upstream.text();
    const matches = html.match(/[A-Z]{4}\d{4}/g) || [];
    const codes = Array.from(new Set(matches));
    res.setHeader('Cache-Control', 'public, s-maxage=300, stale-while-revalidate=600');
    return res.status(200).json({ codes, fetchedAt: new Date().toISOString() });
  } catch (e) {
    return res.status(500).json({ error: e.message || '사이트 fetch 실패' });
  }
}
