# -*- coding: utf-8 -*-
"""
Calculo real do mapa astral (planetas, ascendente, meio do ceu, casas Placidus)
a partir de nome + data + hora + local de nascimento.

Usa Skyfield (efemerides JPL DE421) para as posicoes planetarias e uma
implementacao numerica (bissecao) do sistema de casas Placidus, validada
contra os valores conhecidos do PDF-modelo (Barbara Mendes Silva e Pinho:
Cuiaba/MT, 04/06/1996 21:45 UTC-4 -> Ascendente 16 04' Aquario, MC 16 23'
Escorpiao).
"""
import math
from datetime import datetime

from skyfield.api import load, wgs84
from timezonefinder import TimezoneFinder
from zoneinfo import ZoneInfo
from geopy.geocoders import Nominatim

SIGNS = ["Áries", "Touro", "Gêmeos", "Câncer", "Leão", "Virgem", "Libra",
         "Escorpião", "Sagitário", "Capricórnio", "Aquário", "Peixes"]

_eph = None
_ts = None
_tf = None

def _lazy_init():
    global _eph, _ts, _tf
    if _eph is None:
        _eph = load('de421.bsp')
        _ts = load.timescale()
        _tf = TimezoneFinder()

def sign_deg_min(longitude):
    longitude = longitude % 360
    idx = int(longitude // 30)
    within = longitude - idx * 30
    deg = int(within)
    minute = int(round((within - deg) * 60))
    if minute == 60:
        minute = 0
        deg += 1
        if deg == 30:
            deg = 0
            idx = (idx + 1) % 12
    return SIGNS[idx], deg, minute

def geocode_place(place_text):
    """place_text: 'Cidade, Estado, Pais'. Retorna (lat, lon)."""
    geolocator = Nominatim(user_agent="azimute_mapa_astral_app")
    loc = geolocator.geocode(place_text, timeout=15)
    if loc is None:
        raise ValueError(f"Não foi possível localizar '{place_text}'. Tente um formato como 'Cidade, Estado, País'.")
    return loc.latitude, loc.longitude

def local_to_utc(date_str, time_str, lat, lon):
    """date_str 'YYYY-MM-DD', time_str 'HH:MM'. Usa o fuso horario historico
    correto do local (via timezonefinder + zoneinfo/tzdata)."""
    _lazy_init()
    tzname = _tf.timezone_at(lat=lat, lng=lon)
    if tzname is None:
        raise ValueError("Não foi possível determinar o fuso horário do local informado.")
    naive = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
    local_dt = naive.replace(tzinfo=ZoneInfo(tzname))
    utc_dt = local_dt.astimezone(ZoneInfo("UTC"))
    return utc_dt, tzname

# ---------------------------------------------------------------- astronomy

def _obliquity_deg(t):
    """Obliquidade media da eclitica (Meeus), em graus."""
    T = (t.tt - 2451545.0) / 36525.0
    eps = 23.0 + 26.0/60.0 + 21.448/3600.0
    eps -= (46.8150/3600.0) * T
    eps -= (0.00059/3600.0) * T**2
    eps += (0.001813/3600.0) * T**3
    return eps

def _ra_dec_from_ecliptic_lon(lam_deg, eps_deg, beta_deg=0.0):
    lam = math.radians(lam_deg); eps = math.radians(eps_deg); beta = math.radians(beta_deg)
    sin_dec = math.sin(beta)*math.cos(eps) + math.cos(beta)*math.sin(eps)*math.sin(lam)
    dec = math.asin(max(-1.0, min(1.0, sin_dec)))
    y = math.sin(lam)*math.cos(eps) - math.tan(beta)*math.sin(eps)
    x = math.cos(lam)
    ra = math.degrees(math.atan2(y, x)) % 360
    return ra, math.degrees(dec)

def _ramc_deg(t, lon_east_deg):
    gast_hours = t.gast
    lst_hours = (gast_hours + lon_east_deg / 15.0) % 24
    return lst_hours * 15.0

def _ascendant_mc(ramc_deg, lat_deg, eps_deg):
    ramc = math.radians(ramc_deg); eps = math.radians(eps_deg); lat = math.radians(lat_deg)
    mc = math.degrees(math.atan2(math.sin(ramc), math.cos(ramc)*math.cos(eps))) % 360
    denom = -(math.sin(eps)*math.tan(lat) + math.cos(eps)*math.sin(ramc))
    asc = math.degrees(math.atan2(math.cos(ramc), denom)) % 360
    return asc, mc

def _semi_arc(dec_deg, lat_deg):
    """Arco semi-diurno (graus). None se circumpolar (nao ocorre em latitudes uso comum)."""
    val = -math.tan(math.radians(lat_deg)) * math.tan(math.radians(dec_deg))
    val = max(-1.0, min(1.0, val))
    return math.degrees(math.acos(val))

def _placidus_intermediate(ramc_deg, lat_deg, eps_deg, fraction, quadrant_start_deg, search_span, diurnal):
    """Resolve numericamente (bissecao) a longitude eclitica cuja relacao
    RA(lambda) = ramc + fraction*arco(lambda)  [diurno]  ou
    RA(lambda) = ramc + 180 - fraction*arco(lambda) [noturno, 'diurnal=False']
    e satisfeita, varrendo lambda dentro do quadrante indicado."""
    def f(lam):
        ra, dec = _ra_dec_from_ecliptic_lon(lam, eps_deg)
        try:
            arc = _semi_arc(dec, lat_deg)
        except ValueError:
            arc = 90.0
        if diurnal:
            target = (ramc_deg + fraction * arc) % 360
        else:
            nda = 180 - arc  # arco semi-noturno
            target = (ramc_deg + 180 - fraction * nda) % 360
        diff = (ra - target + 180) % 360 - 180
        return diff

    lo = quadrant_start_deg
    hi = quadrant_start_deg + search_span
    n = 720
    prev_x = lo
    prev_v = f(lo % 360)
    root = None
    for i in range(1, n + 1):
        x = lo + (hi - lo) * i / n
        v = f(x % 360)
        if prev_v == 0:
            root = prev_x
            break
        if (prev_v < 0) != (v < 0):
            a, b = prev_x, x
            fa = prev_v
            for _ in range(60):
                m = (a + b) / 2
                fm = f(m % 360)
                if (fa < 0) == (fm < 0):
                    a, fa = m, fm
                else:
                    b = m
            root = (a + b) / 2
            break
        prev_x, prev_v = x, v
    if root is None:
        root = quadrant_start_deg + search_span / 2
    return root % 360

def compute_houses_placidus(ramc_deg, lat_deg, eps_deg):
    asc, mc = _ascendant_mc(ramc_deg, lat_deg, eps_deg)
    ic = (mc + 180) % 360
    desc = (asc + 180) % 360

    def quadrant_span(a, b):
        span = (b - a) % 360
        return span

    # Cusp 11 fica mais perto do MC (fracao 1/3 do arco MC->ASC); cusp 12 mais
    # perto do ASC (fracao 2/3). Validado contra os valores conhecidos do
    # PDF-modelo (Barbara).
    span_mc_asc = quadrant_span(mc, asc)
    c11 = _placidus_intermediate(ramc_deg, lat_deg, eps_deg, 1/3, mc, span_mc_asc, diurnal=True)
    c12 = _placidus_intermediate(ramc_deg, lat_deg, eps_deg, 2/3, mc, span_mc_asc, diurnal=True)

    # Cusp 2 fica mais perto do ASC (fracao 2/3 do arco noturno); cusp 3 mais
    # perto do IC (fracao 1/3).
    span_ic_desc_rev = quadrant_span(asc, ic)
    c2 = _placidus_intermediate(ramc_deg, lat_deg, eps_deg, 2/3, asc, span_ic_desc_rev, diurnal=False)
    c3 = _placidus_intermediate(ramc_deg, lat_deg, eps_deg, 1/3, asc, span_ic_desc_rev, diurnal=False)

    c5 = (c11 + 180) % 360
    c6 = (c12 + 180) % 360
    c8 = (c2 + 180) % 360
    c9 = (c3 + 180) % 360

    cusps = {1: asc, 2: c2, 3: c3, 4: ic, 5: c5, 6: c6,
             7: desc, 8: c8, 9: c9, 10: mc, 11: c11, 12: c12}
    return cusps, asc, mc

# ------------------------------------------------------------------ planets

PLANET_KEYS = {
    "Sol": ("Identidade, propósito e expressão", "sun"),
    "Lua": ("Segurança emocional e necessidades", "moon"),
    "Mercúrio": ("Pensamento e comunicação", "mercury"),
    "Vênus": ("Afeto, prazer e valores", "venus"),
    "Marte": ("Ação, desejo e iniciativa", "mars"),
    "Júpiter": ("Expansão e confiança", "jupiter barycenter"),
    "Saturno": ("Estrutura, disciplina e dever", "saturn barycenter"),
    "Urano": ("Inovação e liberdade", "uranus barycenter"),
    "Netuno": ("Imaginação e sensibilidade", "neptune barycenter"),
    "Plutão": ("Transformação e poder", "pluto barycenter"),
}

def _mean_lunar_node_lon(t):
    T = (t.tt - 2451545.0) / 36525.0
    omega = 125.04452 - 1934.136261 * T + 0.0020708 * T**2 + T**3 / 450000.0
    return omega % 360

def _which_house(longitude, cusps):
    for h in range(1, 13):
        start = cusps[h]
        end = cusps[h % 12 + 1]
        span = (end - start) % 360
        rel = (longitude - start) % 360
        if rel < span:
            return h
    return 12

def compute_chart(name, date_str, time_str, place_text, lat=None, lon=None):
    """date_str 'YYYY-MM-DD', time_str 'HH:MM' (hora local no local de nascimento).
    Se lat/lon nao forem passados, geocodifica `place_text`."""
    _lazy_init()
    if lat is None or lon is None:
        lat, lon = geocode_place(place_text)

    utc_dt, tzname = local_to_utc(date_str, time_str, lat, lon)
    t = _ts.utc(utc_dt.year, utc_dt.month, utc_dt.day, utc_dt.hour, utc_dt.minute, utc_dt.second)

    eps = _obliquity_deg(t)
    ramc = _ramc_deg(t, lon)
    cusps, asc, mc = compute_houses_placidus(ramc, lat, eps)

    earth = _eph['earth']

    planet_rows = []
    for pname, (key, ephkey) in PLANET_KEYS.items():
        body = _eph[ephkey]
        # geocentrico (nao topocentrico): convencao padrao em astrologia,
        # evita erro de paralaxe na Lua (~1 grau).
        astrometric = earth.at(t).observe(body).apparent()
        _, elon, _ = astrometric.ecliptic_latlon()
        longitude = elon.degrees % 360
        sign, deg, minute = sign_deg_min(longitude)
        house = _which_house(longitude, cusps)
        planet_rows.append({"name": pname, "sign": sign, "deg": deg, "min": minute,
                             "house": f"{house}ª", "key": key, "longitude": longitude})

    node_lon = _mean_lunar_node_lon(t)
    sign, deg, minute = sign_deg_min(node_lon)
    house = _which_house(node_lon, cusps)
    planet_rows.append({"name": "Nodo Norte", "sign": sign, "deg": deg, "min": minute,
                         "house": f"{house}ª", "key": "Direção de desenvolvimento", "longitude": node_lon})

    asc_sign, asc_deg, asc_min = sign_deg_min(asc)
    mc_sign, mc_deg, mc_min = sign_deg_min(mc)
    planet_rows.insert(0, {"name": "Ascendente", "sign": asc_sign, "deg": asc_deg, "min": asc_min,
                            "house": "-", "key": "Presença e início das experiências", "longitude": asc})
    planet_rows.append({"name": "Meio do Céu", "sign": mc_sign, "deg": mc_deg, "min": mc_min,
                         "house": "-", "key": "Vocação e direção pública", "longitude": mc})

    houses_out = []
    for h in range(1, 13):
        s, d, m = sign_deg_min(cusps[h])
        houses_out.append({"n": h, "sign": s, "deg": d, "min": m})

    aspects = compute_aspects(planet_rows)

    return {
        "name": name,
        "birth": {
            "full_name": name,
            "date": date_str,
            "time": time_str,
            "place": place_text,
            "zodiac_system": "Astrologia ocidental tropical",
            "house_system": "Placidus",
            "ascendant_label": f"{asc_deg:02d}°{asc_min:02d}' {asc_sign}",
            "midheaven_label": f"{mc_deg:02d}°{mc_min:02d}' {mc_sign}",
            "timezone": tzname,
            "lat": lat, "lon": lon,
        },
        "ascendant": {"sign": asc_sign, "deg": asc_deg, "min": asc_min},
        "planets": planet_rows,
        "houses": houses_out,
        "aspects": aspects,
        "cusps": cusps,
    }

# ------------------------------------------------------------------ aspects

ASPECTS_DEF = [
    ("conjunção", 0, 8), ("sextil", 60, 6), ("quadratura", 90, 7),
    ("trígono", 120, 7), ("oposição", 180, 8),
]

def compute_aspects(planet_rows):
    usable = [p for p in planet_rows if p["name"] not in ("Ascendente", "Meio do Céu")]
    found = []
    for i in range(len(usable)):
        for j in range(i + 1, len(usable)):
            a, b = usable[i], usable[j]
            diff = abs(a["longitude"] - b["longitude"]) % 360
            if diff > 180:
                diff = 360 - diff
            for asp_name, angle, orb_max in ASPECTS_DEF:
                orb = abs(diff - angle)
                if orb <= orb_max:
                    found.append({
                        "aspecto": f'{a["name"]} {asp_name} {b["name"]}',
                        "a": a["name"], "b": b["name"], "tipo": asp_name,
                        "orbe": f'{orb:.1f}°',
                        "orbe_val": orb,
                        "harmonico": asp_name in ("sextil", "trígono"),
                    })
                    break
    found.sort(key=lambda x: x["orbe_val"])
    return found

def compute_synastry_aspects(planets_a, planets_b):
    """Aspectos ENTRE dois mapas distintos (ex: tutor x pet), em vez de dentro
    de um unico mapa como `compute_aspects`. Compara todo par (planeta de A,
    planeta de B) sem excluir combinacoes repetidas, ja que vem de mapas
    diferentes."""
    usable_a = [p for p in planets_a if p["name"] not in ("Ascendente", "Meio do Céu")]
    usable_b = [p for p in planets_b if p["name"] not in ("Ascendente", "Meio do Céu")]
    found = []
    for a in usable_a:
        for b in usable_b:
            diff = abs(a["longitude"] - b["longitude"]) % 360
            if diff > 180:
                diff = 360 - diff
            for asp_name, angle, orb_max in ASPECTS_DEF:
                orb = abs(diff - angle)
                if orb <= orb_max:
                    found.append({
                        "aspecto": f'{a["name"]} (A) {asp_name} {b["name"]} (B)',
                        "orbe": f'{orb:.1f}°',
                        "orbe_val": orb,
                        "harmonico": asp_name in ("sextil", "trígono"),
                    })
                    break
    found.sort(key=lambda x: x["orbe_val"])
    return found

def compute_house_overlay(planets_a, cusps_b):
    """Para cada planeta de A (ex: tutor), em qual casa do mapa de B (ex: pet)
    ele cai -- e vice-versa, chamando com os argumentos trocados."""
    usable = [p for p in planets_a if p["name"] not in ("Ascendente", "Meio do Céu")]
    return [{"planet": p["name"], "house_in_b": _which_house(p["longitude"], cusps_b)} for p in usable]
