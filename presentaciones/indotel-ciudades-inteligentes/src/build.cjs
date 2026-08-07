const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "INDOTEL";
pres.company = "Instituto Dominicano de las Telecomunicaciones";
pres.title = "Estrategia Nacional de Ciudades Inteligentes";
pres.subject = "Instrumento Estratégico INDOTEL — InDiCo Global / OASC / U4SSC";

/* ─────────────────────────── Sistema de diseño ─────────────────────────── */
const NAVY      = "0B2C5E"; // primario
const NAVY_DEEP = "07203F"; // portada / cierre
const BLUE      = "1668B8"; // círculos de icono, numeración
const SKY       = "2E9BF0"; // acento sobre fondo oscuro
const ICE       = "BBD7F2"; // texto secundario sobre oscuro
const INK       = "0B2C5E"; // titulares sobre claro
const BODY      = "44566B"; // cuerpo
const MUTED     = "8494A5"; // pies y fuentes
const LIGHT     = "F3F7FB"; // tarjetas / fondo suave
const LINE      = "D9E2EC";
const GREEN     = "2E9E5B";

const FONT  = "Calibri";
const MARK  = "Arial"; // wordmark

const W = 13.333, H = 7.5;
const M = 0.62;                 // margen lateral
const CW = W - M * 2;           // ancho útil

const A = (f) => path.join(__dirname, "derived", f);
const I = (n) => path.join(__dirname, "icons", n + ".png");

/* ───────────────────────────── Utilidades ─────────────────────────────── */

// Estimación conservadora de líneas para dimensionar cajas de texto.
function estLines(text, wIn, fontSize) {
  const cpl = Math.max(8, Math.floor((wIn * 72) / (fontSize * 0.50)));
  return text.split("\n").reduce((n, seg) => n + Math.max(1, Math.ceil(seg.length / cpl)), 0);
}
const lineH = (fs) => (fs * 1.28) / 72; // alto de línea en pulgadas

function shadow(opts) { // objeto nuevo cada vez (pptxgenjs muta en sitio)
  return Object.assign({ type: "outer", color: "0B2C5E", opacity: 0.16, blur: 14, offset: 4, angle: 90 }, opts || {});
}

// Cabecera estándar de las láminas de contenido (3–10)
function header(slide, title, subtitle, num, opts = {}) {
  const dark = !!opts.dark;
  slide.addText("INDOTEL", {
    x: M, y: 0.28, w: 4, h: 0.3, fontFace: MARK, fontSize: 15, bold: true,
    color: dark ? "FFFFFF" : NAVY, charSpacing: 1.4, margin: 0, valign: "middle",
  });
  slide.addText("Informe estratégico", {
    x: M, y: 0.585, w: 4, h: 0.24, fontFace: FONT, fontSize: 9.5,
    color: dark ? ICE : MUTED, margin: 0, valign: "middle",
  });
  if (title) {
    slide.addText(title, {
      x: M, y: 0.9, w: opts.titleW || CW, h: opts.titleH || 0.46,
      fontFace: FONT, fontSize: opts.titleSize || 25, bold: true,
      color: dark ? "FFFFFF" : INK, margin: 0, valign: "middle",
    });
  }
  if (subtitle) {
    slide.addText(subtitle, {
      x: M, y: (opts.subY !== undefined ? opts.subY : 1.4), w: opts.subW || CW, h: 0.5,
      fontFace: FONT, fontSize: opts.subSize || 12.5, color: dark ? ICE : BODY,
      margin: 0, valign: "top", lineSpacing: 16,
    });
  }
  if (num) pageNum(slide, num, dark);
}

function pageNum(slide, num, dark) {
  slide.addShape(pres.ShapeType.rect, {
    x: W - 0.62, y: H - 0.58, w: 0.3, h: 0.3, fill: { color: dark ? SKY : NAVY }, line: { type: "none" },
  });
  slide.addText(String(num), {
    x: W - 0.62, y: H - 0.58, w: 0.3, h: 0.3, fontFace: FONT, fontSize: 9.5, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0,
  });
}

function sources(slide, text) {
  slide.addText(text, {
    x: M, y: H - 0.5, w: CW - 0.9, h: 0.26, fontFace: FONT, fontSize: 8.5,
    color: MUTED, italic: true, margin: 0, valign: "middle",
  });
}

// Banda navy inferior con mensaje estratégico
function band(slide, text, y, h, num) {
  slide.addShape(pres.ShapeType.rect, { x: 0, y, w: W, h, fill: { color: NAVY }, line: { type: "none" } });
  slide.addText(text, {
    x: M, y, w: W - M - 1.0, h, fontFace: FONT, fontSize: 11.5, color: "FFFFFF",
    margin: 0, valign: "middle", lineSpacing: 15,
  });
  if (num) {
    slide.addText(String(num), {
      x: W - 0.72, y, w: 0.35, h, fontFace: FONT, fontSize: 9.5, bold: true,
      color: SKY, align: "center", valign: "middle", margin: 0,
    });
  }
}

// Lista con viñeta de icono. Devuelve la Y final.
function iconList(slide, items, o) {
  const fs = o.fontSize || 11;
  const lh = lineH(fs);
  const ind = o.indent || 0.33;
  let y = o.y;
  items.forEach((t) => {
    slide.addImage({ path: I(o.icon), x: o.x, y: y + 0.035, w: 0.18, h: 0.18 });
    const n = estLines(t, o.w - ind, fs);
    slide.addText(t, {
      x: o.x + ind, y: y - 0.045, w: o.w - ind, h: n * lh + 0.12,
      fontFace: FONT, fontSize: fs, color: o.color || BODY, margin: 0, valign: "top", lineSpacing: lh * 72,
    });
    y += n * lh + (o.gap !== undefined ? o.gap : 0.15);
  });
  return y;
}

// Item numerado / con icono: círculo + titular + cuerpo. Devuelve la Y final.
function numItem(slide, o) {
  const d = o.d || 0.44;
  slide.addShape(pres.ShapeType.ellipse, {
    x: o.x, y: o.y, w: d, h: d, fill: { color: o.circle || BLUE }, line: { type: "none" },
  });
  if (o.num !== undefined) {
    slide.addText(String(o.num), {
      x: o.x, y: o.y, w: d, h: d, fontFace: FONT, fontSize: o.numSize || 15, bold: true,
      color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
  } else if (o.icon) {
    const s = d * 0.5;
    slide.addImage({ path: I(o.icon), x: o.x + (d - s) / 2, y: o.y + (d - s) / 2, w: s, h: s });
  }
  const tx = o.x + d + 0.28;
  const tw = o.w - d - 0.28;
  slide.addText(o.title, {
    x: tx, y: o.y - 0.04, w: tw, h: 0.27, fontFace: FONT, fontSize: o.titleSize || 13.5, bold: true,
    color: o.titleColor || INK, margin: 0, valign: "middle",
  });
  const fs = o.bodySize || 10.5;
  const lh = lineH(fs);
  const n = estLines(o.body, tw, fs);
  slide.addText(o.body, {
    x: tx, y: o.y + 0.25, w: tw, h: n * lh + 0.12,
    fontFace: FONT, fontSize: fs, color: o.bodyColor || BODY, margin: 0, valign: "top", lineSpacing: lh * 72,
  });
  return Math.max(o.y + d, o.y + 0.25 + n * lh) + (o.gap !== undefined ? o.gap : 0.24);
}

function vline(slide, x, y, h, color, transparency) {
  slide.addShape(pres.ShapeType.line, {
    x, y, w: 0, h, line: { color, width: 1, transparency: transparency || 0 },
  });
}

/* ══════════════════════════ SLIDE 1 — PORTADA ══════════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY_DEEP };
  s.addImage({ path: A("city_night.jpg"), x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
  // Velo degradado (imagen): un rectángulo semitransparente dejaría una costura vertical.
  s.addImage({ path: A("veil_left.png"), x: 0, y: 0, w: W, h: H });

  s.addText("INDOTEL", {
    x: M, y: 0.42, w: 5, h: 0.5, fontFace: MARK, fontSize: 26, bold: true,
    color: "FFFFFF", charSpacing: 2.2, margin: 0, valign: "middle",
  });

  s.addText("REPÚBLICA DOMINICANA", {
    x: M, y: 1.5, w: 6, h: 0.3, fontFace: FONT, fontSize: 12.5, bold: true,
    color: SKY, charSpacing: 2.6, margin: 0, valign: "middle",
  });
  s.addText("Estrategia Nacional\nde Ciudades Inteligentes", {
    x: M, y: 1.86, w: 8.3, h: 1.35, fontFace: FONT, fontSize: 35, bold: true,
    color: "FFFFFF", margin: 0, valign: "top", lineSpacing: 42,
  });
  s.addText("Instrumento Estratégico INDOTEL", {
    x: M, y: 3.24, w: 7.6, h: 0.45, fontFace: FONT, fontSize: 21,
    color: SKY, margin: 0, valign: "middle",
  });
  s.addText("Conectando municipios. Impulsando innovación.\nTransformando vidas.", {
    x: M, y: 3.88, w: 6.6, h: 0.8, fontFace: FONT, fontSize: 13.5,
    color: "E6EEF7", margin: 0, valign: "top", lineSpacing: 21,
  });

  const pillars = [
    ["s1_datos", "DATOS"],
    ["s1_interop", "INTEROPERABILIDAD"],
    ["s1_estandares", "ESTÁNDARES"],
    ["s1_conectividad", "CONECTIVIDAD"],
    ["s1_sostenibilidad", "SOSTENIBILIDAD"],
  ];
  const colW = 1.66;
  pillars.forEach(([ic, label], i) => {
    const cx = M + i * colW;
    s.addImage({ path: I(ic), x: cx + 0.06, y: 5.32, w: 0.4, h: 0.4 });
    s.addText(label, {
      x: cx, y: 5.86, w: colW - 0.05, h: 0.24, fontFace: FONT, fontSize: 7.5,
      bold: true, color: "D7E5F4", charSpacing: 0.4, margin: 0, valign: "middle",
    });
  });

  // Franja de pie con la identificación del informe
  s.addShape(pres.ShapeType.rect, { x: 0, y: 6.86, w: W, h: 0.64, fill: { color: NAVY }, line: { type: "none" } });
  s.addText([
    { text: "Informe Estratégico", options: { bold: true } },
    { text: "   |   Proyecto InDiCo Global / OASC / U4SSC" },
  ], { x: M, y: 6.86, w: 7.2, h: 0.64, fontFace: FONT, fontSize: 10, color: "FFFFFF", margin: 0, valign: "middle" });
  s.addText([
    { text: "Consejo Directivo INDOTEL" },
    { text: "   |   2026", options: { bold: true } },
  ], { x: 7.9, y: 6.86, w: 4.6, h: 0.64, fontFace: FONT, fontSize: 10, color: "FFFFFF", align: "right", margin: 0, valign: "middle" });
  s.addText("1", {
    x: 12.6, y: 6.86, w: 0.4, h: 0.64, fontFace: FONT, fontSize: 9.5, bold: true,
    color: SKY, align: "center", valign: "middle", margin: 0,
  });

  s.addNotes("Portada del informe estratégico ante el Consejo Directivo del INDOTEL. Los cinco pilares de la franja inferior — datos, interoperabilidad, estándares, conectividad y sostenibilidad — son los ejes que estructuran toda la presentación.");
}

/* ═════════════════════════ SLIDE 2 — LA VISIÓN ═════════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };

  s.addImage({ path: A("monumento.jpg"), x: 6.5, y: 0, w: 6.833, h: 5.05, sizing: { type: "cover", w: 6.833, h: 5.05 } });

  s.addText("1. LA VISIÓN", {
    x: M, y: 0.42, w: 4, h: 0.28, fontFace: FONT, fontSize: 11, bold: true,
    color: NAVY, charSpacing: 1.6, margin: 0, valign: "middle",
  });
  s.addText("Un país conectado.\nMunicipios más eficientes.\nCiudadanos mejor servidos.", {
    x: M, y: 0.85, w: 5.75, h: 1.5, fontFace: FONT, fontSize: 24, bold: true,
    color: INK, margin: 0, valign: "top", lineSpacing: 31,
  });
  s.addText("El instrumento estratégico del INDOTEL convierte la cooperación internacional en resultados reales para la República Dominicana.", {
    x: M, y: 2.52, w: 5.2, h: 1.1, fontFace: FONT, fontSize: 12.5,
    color: BODY, margin: 0, valign: "top", lineSpacing: 18,
  });

  // Tarjeta flotante sobre la fotografía
  s.addShape(pres.ShapeType.roundRect, {
    x: 8.6, y: 0.52, w: 4.28, h: 3.86, rectRadius: 0.1,
    fill: { color: "FFFFFF", transparency: 6 }, line: { type: "none" }, shadow: shadow({ blur: 20, opacity: 0.22 }),
  });
  const benefits = [
    ["s2_servicios", "Mejores servicios\npara la ciudadanía"],
    ["s2_decisiones", "Decisiones basadas\nen datos"],
    ["s2_inversion", "Inversión pública\nmás eficiente"],
    ["s2_ciudades", "Ciudades sostenibles\ne inclusivas"],
  ];
  benefits.forEach(([ic, label], i) => {
    const y = 0.86 + i * 0.9;
    s.addShape(pres.ShapeType.ellipse, { x: 8.95, y, w: 0.5, h: 0.5, fill: { color: BLUE }, line: { type: "none" } });
    s.addImage({ path: I(ic), x: 9.07, y: y + 0.12, w: 0.26, h: 0.26 });
    s.addText(label, {
      x: 9.66, y: y - 0.04, w: 3.0, h: 0.58, fontFace: FONT, fontSize: 11,
      color: INK, margin: 0, valign: "middle", lineSpacing: 14.5,
    });
  });

  // Banda de indicadores
  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.05, w: W, h: 2.45, fill: { color: NAVY }, line: { type: "none" } });
  const kpis = [
    ["s2_kpi_municipios", "MUNICIPIOS\nCONECTADOS", "> 50"],
    ["s2_kpi_ciudadanos", "CIUDADANOS\nIMPACTADOS", "> 3M"],
    ["s2_kpi_proyectos", "PROYECTOS\nPRIORIZADOS", "> 100"],
    ["s2_kpi_decision", "MEJOR DECISIÓN\nPÚBLICA", "100%"],
  ];
  const kw = W / 4;
  kpis.forEach(([ic, label, val], i) => {
    const bx = i * kw;
    s.addImage({ path: I(ic), x: bx + 0.5, y: 5.86, w: 0.62, h: 0.62 });
    s.addText(label, {
      x: bx + 1.32, y: 5.5, w: kw - 1.55, h: 0.55, fontFace: FONT, fontSize: 10, bold: true,
      color: "FFFFFF", charSpacing: 0.8, margin: 0, valign: "middle", lineSpacing: 13,
    });
    s.addText(val, {
      x: bx + 1.32, y: 6.12, w: kw - 1.55, h: 0.62, fontFace: FONT, fontSize: 27, bold: true,
      color: "FFFFFF", margin: 0, valign: "middle",
    });
    if (i > 0) vline(s, bx, 5.5, 1.5, "FFFFFF", 72);
  });

  s.addNotes("La visión ordena el discurso: conectividad, eficiencia municipal y mejor servicio al ciudadano. Los cuatro indicadores de la banda inferior son las metas de referencia del instrumento estratégico.");
}

/* ══════════════════ SLIDE 3 — QUIÉNES SON LOS ACTORES ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "Primero: Quiénes son los actores", null, null);

  s.addImage({ path: A("globe_crop.jpg"), x: 0, y: 1.5, w: 6.05, h: 4.45, sizing: { type: "cover", w: 6.05, h: 4.45 } });

  let y = 1.78;
  const actores = [
    [1, "InDiCo Global", "Proyecto de cooperación internacional sobre estándares digitales/ICT y políticas digitales. Su valor para RD es transferencia de capacidades, alineación con estándares y apoyo técnico."],
    [2, "OASC", "Open & Agile Smart Cities and Communities. Red global de ciudades y comunidades que ayuda a administraciones locales a avanzar en transformación digital e interoperabilidad."],
    [3, "U4SSC", "United for Smart Sustainable Cities. Iniciativa del sistema ONU que ofrece KPIs y modelos de medición para ciudades inteligentes y sostenibles."],
  ];
  actores.forEach(([n, t, b]) => {
    y = numItem(s, { x: 6.4, y, w: 6.3, num: n, title: t, body: b, gap: 0.52 });
  });

  band(s, "Resumen estratégico: InDiCo abre la cooperación; OASC aporta interoperabilidad y red; U4SSC aporta medición, INDOTEL convierte todo en política pública operativa.", 6.15, 0.72, 3);
  sources(s, "Fuentes: InDiCo Global; OASC; U4SSC / ITU.");

  s.addNotes("Tres actores, tres aportes distintos y complementarios. Es la lámina que evita la confusión más común: creer que InDiCo, OASC y U4SSC son lo mismo.");
}

/* ═══════════════════════ SLIDE 4 — InDiCo GLOBAL ═══════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "InDiCo Global: cooperación digital y estándares",
    "No es solo un “proyecto”; es una puerta de acceso a conocimiento y estándares internacionales.", null);

  const cw = 5.83, gap = 0.43;
  const cols = [
    ["Qué promueve", 0, [
      "Conocimiento sobre modelos de estándares ICT/digitales.",
      "Capacitación y cooperación con países y regiones socias.",
      "Interoperabilidad, ciberseguridad, datos, IoT, 5G, IA y ciudades inteligentes.",
      "Uso de estándares compatibles con valores europeos: apertura, protección de datos y confianza.",
    ]],
    ["Aplicación directa en RD", 1, [
      "OASC trabajará con 3–5 ciudades dominicanas para apoyar su madurez digital.",
      "Se usarán LORDIMAS y el modelo ITU de madurez de ciudades inteligentes y sostenibles.",
      "Meta publicada: crear un hub sostenible vinculado a OASC y U4SSC.",
    ]],
  ];
  cols.forEach(([title, i, items]) => {
    const cx = M + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 2.12, w: cw, h: 3.62, rectRadius: 0.06, fill: { color: LIGHT }, line: { type: "none" },
    });
    s.addText(title, {
      x: cx + 0.4, y: 2.42, w: cw - 0.8, h: 0.32, fontFace: FONT, fontSize: 13.5, bold: true,
      color: BLUE, margin: 0, valign: "middle",
    });
    iconList(s, items, { x: cx + 0.4, y: 3.02, w: cw - 0.8, icon: "check_blue", fontSize: 11, gap: 0.30 });
  });

  band(s, "Para INDOTEL: esto permite pasar de una cooperación puntual a una arquitectura nacional de acompañamiento municipal.", 6.06, 0.72, 4);
  sources(s, "Fuentes: InDiCo Global; InDiCo-Global Second Open Call Winners, 2026.");

  s.addNotes("InDiCo Global no aporta solo financiamiento: aporta acceso a estándares, capacitación y una red de pares. La aplicación en RD ya tiene metas concretas — de 3 a 5 ciudades piloto y un hub sostenible.");
}

/* ═══════════════════════════ SLIDE 5 — OASC ════════════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "OASC: Ciudades Inteligentes Abiertas y Ágiles",
    "Open & Agile Smart Cities and Communities es una red global de ciudades y comunidades.", null);

  const cards = [
    ["s5_naturaleza", BLUE, "Naturaleza", "Organización internacional con sede en Bruselas que conecta ciudades, regiones y comunidades para compartir soluciones digitales basadas en datos."],
    ["s5_tecnico", NAVY, "Enfoque técnico", "Trabaja la interoperabilidad técnica, semántica, organizacional y legal, para que sistemas y datos puedan combinarse y reutilizarse."],
    ["s5_valor", "0E7A6B", "Valor público", "Busca evitar dependencia de proveedores, reducir costos, mejorar eficiencia y facilitar la réplica de soluciones entre ciudades."],
  ];
  const cw = 3.86, gap = 0.375;
  cards.forEach(([ic, col, title, body], i) => {
    const cx = M + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 2.15, w: cw, h: 2.92, rectRadius: 0.07,
      fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 }, shadow: shadow({ blur: 12, opacity: 0.1, offset: 3 }),
    });
    s.addShape(pres.ShapeType.ellipse, { x: cx + cw / 2 - 0.35, y: 2.46, w: 0.7, h: 0.7, fill: { color: col }, line: { type: "none" } });
    s.addImage({ path: I(ic), x: cx + cw / 2 - 0.18, y: 2.63, w: 0.36, h: 0.36 });
    s.addText(title, {
      x: cx + 0.25, y: 3.33, w: cw - 0.5, h: 0.32, fontFace: FONT, fontSize: 14.5, bold: true,
      color: INK, align: "center", margin: 0, valign: "middle",
    });
    s.addText(body, {
      x: cx + 0.34, y: 3.72, w: cw - 0.68, h: 1.28, fontFace: FONT, fontSize: 11,
      color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 15.5,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.52, w: CW, h: 0.9, rectRadius: 0.06, fill: { color: "E8F1FA" }, line: { type: "none" },
  });
  s.addText("OASC no vende una plataforma cerrada; impulsa reglas comunes para que las plataformas municipales puedan comunicarse.", {
    x: M + 0.42, y: 5.52, w: CW - 0.84, h: 0.9, fontFace: FONT, fontSize: 12.5,
    color: NAVY, align: "center", margin: 0, valign: "middle", lineSpacing: 17,
  });

  sources(s, "Fuentes: OASC official site; Living-in.eu OASC partner profile.");
  pageNum(s, 5);

  s.addNotes("El punto clave a comunicar: OASC no es un proveedor. Es una red que define reglas comunes. Eso protege al Estado de la dependencia tecnológica.");
}

/* ═══════════════════════ SLIDE 6 — CÓMO OPERA OASC ═════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "Cómo opera OASC",
    "La organización combina gobernanza de ciudades, soporte técnico y participación en proyectos internacionales.",
    6, { subW: 6.6 });

  s.addImage({ path: A("rooftop.jpg"), x: 7.28, y: 1.50, w: 5.44, h: 4.62, sizing: { type: "cover", w: 5.44, h: 4.62 } });

  let y = 2.16;
  const ops = [
    ["s6_gobernanza", "Gobernanza", "Consejo de Ciudades, órgano de gobierno y Consejo Tecnológico para validar el rumbo y la adopción de MIMs."],
    ["s6_servicios", "Servicios", "Conocimiento, proyectos, herramientas técnicas, apoyo a miembros, capacitación y espacios de colaboración."],
    ["s6_proyectos", "Proyectos", "Participa en iniciativas de datos, gemelos digitales, CitiVerse, ciberseguridad, movilidad y ciudades inteligentes."],
    ["s6_itu", "Relación con ITU", "En 2024 OASC y la UIT firmaron cooperación para llevar la voz de ciudades y comunidades a tecnologías emergentes."],
  ];
  ops.forEach(([ic, t, b]) => {
    y = numItem(s, { x: M, y, w: 6.3, icon: ic, title: t, body: b, gap: 0.45, titleSize: 13 });
  });

  s.addNotes("Cuatro mecanismos de operación. El acuerdo con la UIT de 2024 es el que da respaldo institucional al vínculo con el sistema de Naciones Unidas.");
}

/* ═════════════════ SLIDE 7 — CONCEPTO CENTRAL: MIMs ════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "El concepto central de OASC: Mecanismos Mínimos de Interoperabilidad",
    "Minimal Interoperability Mechanisms", null, { titleSize: 21, titleH: 0.95, subY: 1.76, titleW: 11.2 });

  const cw = 5.83, gap = 0.43;
  const cols = [
    ["¿Qué resuelven?", 0, [
      "Permiten un nivel mínimo y suficiente de interoperabilidad entre datos, sistemas y servicios.",
      "Combinan soluciones técnicas y no técnicas.",
      "Ayudan a que las inversiones digitales sean más futuras y menos dependientes de un proveedor.",
      "Mejoran la cooperación entre departamentos y servicios públicos.",
    ]],
    ["Interpretación para INDOTEL", 1, [
      "Un municipio no debe comprar sistemas aislados.",
      "Los datos de movilidad, ambiente, permisos, seguridad y conectividad deben poder integrarse.",
      "INDOTEL puede promover lenguaje común, guías técnicas y criterios mínimos para proyectos municipales.",
    ]],
  ];
  cols.forEach(([title, i, items]) => {
    const cx = M + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 2.15, w: cw, h: 3.45, rectRadius: 0.06, fill: { color: LIGHT }, line: { type: "none" },
    });
    s.addText(title, {
      x: cx + 0.4, y: 2.44, w: cw - 0.8, h: 0.32, fontFace: FONT, fontSize: 13.5, bold: true,
      color: BLUE, margin: 0, valign: "middle",
    });
    iconList(s, items, { x: cx + 0.4, y: 2.95, w: cw - 0.8, icon: "check_blue", fontSize: 11, gap: 0.26 });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.82, w: CW, h: 0.85, rectRadius: 0.06, fill: { color: "E8F1FA" }, line: { type: "none" },
  });
  s.addImage({ path: I("quote_shield"), x: M + 0.42, y: 6.10, w: 0.3, h: 0.3 });
  s.addText("“Los MIMs convierten la ciudad inteligente en gestión pública más eficiente, no en tecnología aislada.”", {
    x: M + 0.92, y: 5.82, w: CW - 1.35, h: 0.85, fontFace: FONT, fontSize: 12.5, italic: true,
    color: NAVY, margin: 0, valign: "middle", lineSpacing: 17,
  });

  sources(s, "Fuentes: OASC MIMs concept and materials.");
  pageNum(s, 7);

  s.addNotes("Traducción del concepto técnico a decisión de política pública: el mínimo común obligatorio evita compras aisladas y protege la inversión municipal.");
}

/* ══════════════════════ SLIDE 8 — ¿CÓMO OPERA OASC? ════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "¿Cómo opera OASC?", "Gobernanza, conocimiento y proyectos para ciudades más inteligentes.", 8);

  const pillars = [
    ["s8_gobernanza", "GOBERNANZA", "Consejo de Ciudades y Consejo Tecnológico para validar el rumbo y estándares (MIMs)."],
    ["s8_servicios", "SERVICIOS", "Conocimiento, herramientas técnicas, capacitación y apoyo continuo a municipios."],
    ["s8_proyectos", "PROYECTOS", "Iniciativas en datos, gemelos digitales, movilidad, ciberseguridad, CitiVerse y más."],
    ["s8_alianzas", "ALIANZAS", "Cooperación con ITU y actores globales para llevar la voz de las ciudades a tecnologías emergentes."],
  ];
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 2.35, w: CW, h: 3.85, rectRadius: 0.07, fill: { color: LIGHT }, line: { type: "none" },
  });
  const cw = CW / 4;
  pillars.forEach(([ic, title, body], i) => {
    const cx = M + i * cw;
    s.addImage({ path: I(ic), x: cx + cw / 2 - 0.5, y: 2.9, w: 1.0, h: 1.0 });
    s.addText(title, {
      x: cx + 0.12, y: 4.18, w: cw - 0.24, h: 0.32, fontFace: FONT, fontSize: 13, bold: true,
      color: INK, align: "center", charSpacing: 0.9, margin: 0, valign: "middle",
    });
    s.addText(body, {
      x: cx + 0.3, y: 4.62, w: cw - 0.6, h: 1.4, fontFace: FONT, fontSize: 11.5,
      color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 16,
    });
    if (i > 0) vline(s, cx, 2.82, 2.95, "C9D8E6");
  });

  s.addNotes("Versión sintética de la lámina 6, pensada para audiencias ejecutivas: cuatro pilares, una línea cada uno.");
}

/* ══════════════════════ SLIDE 9 — CONCEPTO CLAVE: MIMs ═════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "El concepto clave: MIMs",
    "Mecanismos Mínimos de Interoperabilidad. La base para ciudades que se entienden.", 9, { subW: 8 });

  s.addImage({ path: A("mims.jpg"), x: M, y: 1.9, w: 7.55, h: 4.55, sizing: { type: "cover", w: 7.55, h: 4.55 } });

  iconList(s, [
    "Nivel mínimo y suficiente de interoperabilidad.",
    "Soluciones técnicas y no técnicas.",
    "Inversiones más futuras y menos dependientes.",
    "Lenguaje común, guías y criterios mínimos para todos los municipios.",
  ], { x: 8.5, y: 2.62, w: 4.2, icon: "check_green", fontSize: 12, gap: 0.58, indent: 0.36 });

  s.addNotes("Lámina de anclaje visual: los MIMs conectan movilidad, ambiente, permisos, seguridad y conectividad en un mismo lenguaje.");
}

/* ═══════════════════════ SLIDE 10 — HOJA DE RUTA ═══════════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  header(s, "Hoja de ruta INDOTEL", "De la cooperación a la transformación nacional.", 10);

  const steps = [
    ["DIAGNÓSTICO", "Evaluamos madurez digital municipal con LORDIMAS e ITU."],
    ["ESTÁNDARES", "Definimos MIMs, guías y arquitectura nacional."],
    ["CAPACIDADES", "Formamos equipos y acompañamos a municipios."],
    ["PROYECTOS", "Desplegamos casos prioritarios con impacto rápido."],
    ["MEDICIÓN", "Medimos resultados con KPIs U4SSC y mejora continua."],
  ];
  const tones = ["0B2C5E", "0F4382", "1257A0", "1668B8", "2E9BF0"];
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 2.5, w: CW, h: 3.6, rectRadius: 0.07, fill: { color: LIGHT }, line: { type: "none" },
  });
  const cw = CW / 5;
  const cy = 3.55;  // centro vertical de los círculos
  const d = 0.94;

  steps.forEach(([title, body], i) => {
    const cx = M + i * cw + cw / 2;
    if (i < steps.length - 1) {
      s.addShape(pres.ShapeType.line, {
        x: cx + d / 2, y: cy, w: cw - d, h: 0,
        line: { color: "AEC2D6", width: 1.9, endArrowType: "triangle" },
      });
    }
    s.addShape(pres.ShapeType.ellipse, {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: tones[i] }, line: { type: "none" },
    });
    s.addText(String(i + 1), {
      x: cx - d / 2, y: cy - d / 2, w: d, h: d, fontFace: FONT, fontSize: 27, bold: true,
      color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
    s.addText(title, {
      x: cx - cw / 2 + 0.1, y: 4.26, w: cw - 0.2, h: 0.3, fontFace: FONT, fontSize: 13, bold: true,
      color: INK, align: "center", charSpacing: 0.7, margin: 0, valign: "middle",
    });
    s.addText(body, {
      x: cx - cw / 2 + 0.22, y: 4.7, w: cw - 0.44, h: 1.3, fontFace: FONT, fontSize: 11.5,
      color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 16,
    });
  });

  s.addNotes("Cinco fases secuenciales. La medición cierra el ciclo y lo realimenta: sin KPIs U4SSC no hay mejora continua ni evidencia para el Consejo Directivo.");
}

/* ═════════════════ SLIDE 11 — IMPACTO PARA LA CIUDADANÍA ═══════════════ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY_DEEP };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: NAVY_DEEP }, line: { type: "none" } });
  s.addImage({ path: A("couple_opt.png"), x: 7.7, y: 0.62, w: 5.05, h: 4.7 });

  header(s, null, null, null, { dark: true });
  s.addText("Impacto para la ciudadanía", {
    x: M, y: 1.72, w: 6.4, h: 0.6, fontFace: FONT, fontSize: 26, bold: true,
    color: "FFFFFF", margin: 0, valign: "middle",
  });
  s.addText("Mejores servicios. Mejor calidad de vida.", {
    x: M, y: 2.46, w: 6.4, h: 0.45, fontFace: FONT, fontSize: 18,
    color: SKY, margin: 0, valign: "middle",
  });
  s.addText("El instrumento estratégico se mide donde importa: en el servicio diario que recibe cada ciudadano dominicano.", {
    x: M, y: 3.2, w: 5.7, h: 0.9, fontFace: FONT, fontSize: 12.5,
    color: ICE, margin: 0, valign: "top", lineSpacing: 18,
  });

  const impacts = [
    ["s11_salud", "Salud\ndigital"],
    ["s11_movilidad", "Movilidad\ninteligente"],
    ["s11_seguridad", "Seguridad\nciudadana"],
    ["s11_alumbrado", "Alumbrado\neficiente"],
    ["s11_ambiente", "Medio ambiente\nsostenible"],
    ["s11_digital", "Servicios\ndigitales 24/7"],
  ];
  const iw = CW / 6;
  impacts.forEach(([ic, label], i) => {
    const cx = M + i * iw;
    s.addImage({ path: I(ic), x: cx + iw / 2 - 0.25, y: 5.82, w: 0.5, h: 0.5 });
    s.addText(label, {
      x: cx + 0.06, y: 6.48, w: iw - 0.12, h: 0.6, fontFace: FONT, fontSize: 10.5,
      color: "E6EEF7", align: "center", margin: 0, valign: "top", lineSpacing: 14,
    });
    if (i > 0) vline(s, cx, 5.8, 1.25, "FFFFFF", 78);
  });

  pageNum(s, 11, true);
  s.addNotes("Cierre del argumento técnico con el argumento humano: seis servicios cotidianos donde el ciudadano percibe el resultado de la interoperabilidad.");
}

/* ══════════════════════ SLIDE 12 — DECISIÓN ESTRATÉGICA ════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY_DEEP };
  s.addImage({ path: A("city_dark.jpg"), x: 0, y: 2.5, w: W, h: 5.0, sizing: { type: "cover", w: W, h: 5.0 } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: NAVY_DEEP, transparency: 16 }, line: { type: "none" } });

  header(s, null, null, null, { dark: true });

  s.addText("DECISIÓN ESTRATÉGICA", {
    x: M, y: 1.55, w: 6, h: 0.3, fontFace: FONT, fontSize: 11, bold: true,
    color: SKY, charSpacing: 2.2, margin: 0, valign: "middle",
  });
  s.addText("Autorización para convertir esta cooperación en un instrumento permanente del INDOTEL.", {
    x: M, y: 1.98, w: 8.05, h: 1.85, fontFace: FONT, fontSize: 31, bold: true,
    color: "FFFFFF", margin: 0, valign: "top", lineSpacing: 39,
  });
  s.addText("Más ciudades conectadas. Más datos que generan valor.\nMejores decisiones. Un país que avanza.", {
    x: M, y: 4.0, w: 7.6, h: 0.95, fontFace: FONT, fontSize: 15,
    color: ICE, margin: 0, valign: "top", lineSpacing: 24,
  });

  vline(s, 9.42, 1.65, 4.3, "FFFFFF", 70);

  s.addText("INDOTEL", {
    x: 9.95, y: 1.95, w: 3.1, h: 0.55, fontFace: MARK, fontSize: 30, bold: true,
    color: "FFFFFF", charSpacing: 1.8, margin: 0, valign: "middle",
  });
  s.addText("Lideramos la transformación digital de los municipios dominicanos.", {
    x: 9.95, y: 2.68, w: 2.95, h: 1.1, fontFace: FONT, fontSize: 12.5,
    color: ICE, margin: 0, valign: "top", lineSpacing: 18,
  });
  s.addText("2026", {
    x: 9.95, y: 4.55, w: 3.1, h: 0.7, fontFace: FONT, fontSize: 34, bold: true,
    color: SKY, margin: 0, valign: "middle",
  });

  pageNum(s, 12, true);
  s.addNotes("Lámina de cierre y de solicitud formal: el Consejo Directivo autoriza convertir la cooperación InDiCo/OASC/U4SSC en un instrumento permanente del INDOTEL.");
}

const OUT = path.join(__dirname, "..", "INDOTEL-Estrategia-Nacional-Ciudades-Inteligentes.pptx");
pres.writeFile({ fileName: OUT }).then(() => console.log("OK ->", OUT));
