const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "INDOTEL";
pres.company = "Instituto Dominicano de las Telecomunicaciones";
pres.title = "Estrategia Nacional de Ciudades Inteligentes";
pres.subject = "Instrumento Estratégico INDOTEL — InDiCo Global / OASC / U4SSC";

/* ── Sistema de diseño ─────────────────────────────────────────────────────
   Colores muestreados directamente del mosaico de referencia.             */
const NAVY  = "0A2E7A"; // bandas y pie
const BLUE  = "0B4489"; // discos de icono
const INK   = "0D1B5E"; // titulares
const SUB   = "243C8F"; // subtítulos y encabezados de columna
const DEEP  = "04173F"; // láminas oscuras
const SKY   = "3AA0F5"; // acento sobre oscuro
const ICE   = "C3D6EE";
const BODY  = "3B4356";
const MUTED = "8A93A5";
const BG    = "F2F3F8"; // fondo de láminas claras
const RULE  = "CBD5E4"; // divisores finos
const GREEN = "2E9E5B";

const FONT = "Calibri";
const MARK = "Arial";

const W = 13.333, H = 7.5;
const M = 0.62;
const CW = W - M * 2;

const A = (f) => path.join(__dirname, "derived", f);
const I = (n) => path.join(__dirname, "icons", n + ".png");

/* ── Utilidades ──────────────────────────────────────────────────────────── */

function estLines(text, wIn, fontSize) {
  const cpl = Math.max(8, Math.floor((wIn * 72) / (fontSize * 0.50)));
  return text.split("\n").reduce((n, s) => n + Math.max(1, Math.ceil(s.length / cpl)), 0);
}
const lineH = (fs) => (fs * 1.28) / 72;

function shadow(o) { // objeto nuevo por llamada: pptxgenjs muta en sitio
  return Object.assign({ type: "outer", color: "0A1F45", opacity: 0.18, blur: 16, offset: 4, angle: 90 }, o || {});
}

function header(slide, title, subtitle, num, opts = {}) {
  const dark = !!opts.dark;
  slide.addText("INDOTEL", {
    x: M, y: 0.26, w: 4, h: 0.3, fontFace: MARK, fontSize: 15, bold: true,
    color: dark ? "FFFFFF" : INK, charSpacing: 1.3, margin: 0, valign: "middle",
  });
  slide.addText("Informe estratégico", {
    x: M, y: 0.565, w: 4, h: 0.24, fontFace: FONT, fontSize: 9.5,
    color: dark ? ICE : MUTED, margin: 0, valign: "middle",
  });
  if (title) slide.addText(title, {
    x: M, y: opts.titleY !== undefined ? opts.titleY : 0.84, w: opts.titleW || CW,
    h: opts.titleH || 0.46, fontFace: FONT, fontSize: opts.titleSize || 24.5, bold: true,
    color: dark ? "FFFFFF" : INK, margin: 0, valign: "middle",
  });
  // En el original el subtítulo va en azul, no en gris.
  if (subtitle) slide.addText(subtitle, {
    x: M, y: opts.subY !== undefined ? opts.subY : 1.35, w: opts.subW || CW, h: 0.52,
    fontFace: FONT, fontSize: opts.subSize || 12.5, bold: true,
    color: dark ? ICE : SUB, margin: 0, valign: "top", lineSpacing: 16.5,
  });
  if (num) pageNum(slide, num, dark);
}

function pageNum(slide, num, dark) {
  slide.addShape(pres.ShapeType.rect, {
    x: W - 0.58, y: H - 0.54, w: 0.34, h: 0.34,
    fill: { color: dark ? "1268C4" : NAVY }, line: { type: "none" },
  });
  slide.addText(String(num), {
    x: W - 0.58, y: H - 0.54, w: 0.34, h: 0.34, fontFace: FONT, fontSize: 10, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0,
  });
}

function sources(slide, text) {
  slide.addText(text, {
    x: M, y: H - 0.46, w: CW - 1.0, h: 0.26, fontFace: FONT, fontSize: 8.5,
    color: MUTED, margin: 0, valign: "middle",
  });
}

function band(slide, text, y, h, num) {
  slide.addShape(pres.ShapeType.rect, { x: 0, y, w: W, h, fill: { color: NAVY }, line: { type: "none" } });
  slide.addText(text, {
    x: M, y, w: W - M - 1.15, h, fontFace: FONT, fontSize: 11.5, color: "FFFFFF",
    margin: 0, valign: "middle", lineSpacing: 15.5,
  });
  if (num) slide.addText(String(num), {
    x: W - 0.86, y, w: 0.4, h, fontFace: FONT, fontSize: 10, bold: true,
    color: "9CC4EE", align: "center", valign: "middle", margin: 0,
  });
}

function vline(slide, x, y, h, color, transparency) {
  slide.addShape(pres.ShapeType.line, {
    x, y, w: 0, h, line: { color, width: 1, transparency: transparency || 0 },
  });
}

function iconList(slide, items, o) {
  const fs = o.fontSize || 11;
  const lh = lineH(fs);
  const ind = o.indent || 0.34;
  let y = o.y;
  items.forEach((t) => {
    slide.addImage({ path: I(o.icon), x: o.x, y: y + 0.03, w: 0.2, h: 0.2 });
    const n = estLines(t, o.w - ind, fs);
    slide.addText(t, {
      x: o.x + ind, y: y - 0.05, w: o.w - ind, h: n * lh + 0.12,
      fontFace: FONT, fontSize: fs, color: o.color || BODY, margin: 0, valign: "top", lineSpacing: lh * 72,
    });
    y += n * lh + (o.gap !== undefined ? o.gap : 0.16);
  });
  return y;
}

function numItem(slide, o) {
  const d = o.d || 0.46;
  const shp = o.square ? pres.ShapeType.roundRect : pres.ShapeType.ellipse;
  const extra = o.square ? { rectRadius: 0.11 } : {};
  slide.addShape(shp, Object.assign({
    x: o.x, y: o.y, w: d, h: d, fill: { color: o.circle || BLUE }, line: { type: "none" },
  }, extra));
  if (o.num !== undefined) {
    slide.addText(String(o.num), {
      x: o.x, y: o.y, w: d, h: d, fontFace: FONT, fontSize: o.numSize || 15.5, bold: true,
      color: "FFFFFF", align: "center", valign: "middle", margin: 0,
    });
  } else if (o.icon) {
    const s = d * 0.54;
    slide.addImage({ path: I(o.icon), x: o.x + (d - s) / 2, y: o.y + (d - s) / 2, w: s, h: s });
  }
  const tx = o.x + d + 0.3, tw = o.w - d - 0.3;
  slide.addText(o.title, {
    x: tx, y: o.y - 0.05, w: tw, h: 0.28, fontFace: FONT, fontSize: o.titleSize || 13.5, bold: true,
    color: INK, margin: 0, valign: "middle",
  });
  const fs = o.bodySize || 10.5, lh = lineH(fs);
  const n = estLines(o.body, tw, fs);
  slide.addText(o.body, {
    x: tx, y: o.y + 0.25, w: tw, h: n * lh + 0.12,
    fontFace: FONT, fontSize: fs, color: BODY, margin: 0, valign: "top", lineSpacing: lh * 72,
  });
  return { end: Math.max(o.y + d, o.y + 0.25 + n * lh), next: Math.max(o.y + d, o.y + 0.25 + n * lh) + (o.gap !== undefined ? o.gap : 0.26) };
}

/* ══════════════════ 1 · PORTADA ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: DEEP };
  s.addImage({ path: A("city_night.jpg"), x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
  s.addImage({ path: A("veil_cover.png"), x: 0, y: 0, w: W, h: H });

  s.addText("INDOTEL", {
    x: M, y: 0.42, w: 5, h: 0.5, fontFace: MARK, fontSize: 26, bold: true,
    color: "FFFFFF", charSpacing: 2.2, margin: 0, valign: "middle",
  });
  s.addText("REPÚBLICA DOMINICANA", {
    x: M, y: 1.52, w: 6, h: 0.3, fontFace: FONT, fontSize: 12.5, bold: true,
    color: "4FA8F0", charSpacing: 2.6, margin: 0, valign: "middle",
  });
  s.addText("Estrategia Nacional\nde Ciudades Inteligentes", {
    x: M, y: 1.88, w: 8.3, h: 1.35, fontFace: FONT, fontSize: 35, bold: true,
    color: "FFFFFF", margin: 0, valign: "top", lineSpacing: 42,
  });
  s.addText("Instrumento Estratégico INDOTEL", {
    x: M, y: 3.26, w: 7.6, h: 0.45, fontFace: FONT, fontSize: 21,
    color: "4FA8F0", margin: 0, valign: "middle",
  });
  s.addText("Conectando municipios. Impulsando innovación.\nTransformando vidas.", {
    x: M, y: 3.9, w: 6.6, h: 0.8, fontFace: FONT, fontSize: 13.5,
    color: "E4EDF8", margin: 0, valign: "top", lineSpacing: 21,
  });

  [["s1_datos", "DATOS"], ["s1_interop", "INTEROPERABILIDAD"], ["s1_estandares", "ESTÁNDARES"],
   ["s1_conectividad", "CONECTIVIDAD"], ["s1_sostenibilidad", "SOSTENIBILIDAD"]]
  .forEach(([ic, label], i) => {
    const cx = M + i * 1.66;
    s.addImage({ path: I(ic), x: cx + 0.04, y: 5.3, w: 0.44, h: 0.44 });
    s.addText(label, {
      x: cx, y: 5.9, w: 1.61, h: 0.24, fontFace: FONT, fontSize: 7.5, bold: true,
      color: "D3E3F5", charSpacing: 0.4, margin: 0, valign: "middle",
    });
  });

  s.addShape(pres.ShapeType.rect, { x: 0, y: 6.86, w: W, h: 0.64, fill: { color: NAVY }, line: { type: "none" } });
  s.addText([{ text: "Informe Estratégico", options: { bold: true } },
             { text: "   |   Proyecto InDiCo Global / OASC / U4SSC" }], {
    x: M, y: 6.86, w: 7.2, h: 0.64, fontFace: FONT, fontSize: 10, color: "FFFFFF", margin: 0, valign: "middle" });
  s.addText([{ text: "Consejo Directivo INDOTEL" }, { text: "   |   2026", options: { bold: true } }], {
    x: 7.9, y: 6.86, w: 4.5, h: 0.64, fontFace: FONT, fontSize: 10, color: "FFFFFF",
    align: "right", margin: 0, valign: "middle" });
  s.addText("1", { x: 12.55, y: 6.86, w: 0.4, h: 0.64, fontFace: FONT, fontSize: 10, bold: true,
    color: "9CC4EE", align: "center", valign: "middle", margin: 0 });

  s.addNotes("Portada del informe estratégico ante el Consejo Directivo. Los cinco pilares de la franja inferior — datos, interoperabilidad, estándares, conectividad y sostenibilidad — estructuran toda la presentación.");
}

/* ══════════════════ 2 · LA VISIÓN ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addImage({ path: A("monumento_soft.jpg"), x: 5.9, y: 0, w: 7.433, h: 5.12 });

  s.addText("1. LA VISIÓN", {
    x: M, y: 0.42, w: 4, h: 0.28, fontFace: FONT, fontSize: 11, bold: true,
    color: INK, charSpacing: 1.6, margin: 0, valign: "middle" });
  s.addText("Un país conectado.\nMunicipios más eficientes.\nCiudadanos mejor servidos.", {
    x: M, y: 0.86, w: 5.5, h: 1.5, fontFace: FONT, fontSize: 22.5, bold: true,
    color: INK, margin: 0, valign: "top", lineSpacing: 30 });
  s.addText("El instrumento estratégico del INDOTEL convierte la cooperación internacional en resultados reales para la República Dominicana.", {
    x: M, y: 2.62, w: 4.75, h: 1.1, fontFace: FONT, fontSize: 12.5,
    color: BODY, margin: 0, valign: "top", lineSpacing: 18.5 });

  s.addShape(pres.ShapeType.roundRect, {
    x: 8.7, y: 0.5, w: 4.32, h: 3.95, rectRadius: 0.11,
    fill: { color: "FFFFFF", transparency: 5 }, line: { type: "none" }, shadow: shadow({ blur: 22, opacity: 0.24 }) });
  [["s2_servicios", "Mejores servicios\npara la ciudadanía"],
   ["s2_decisiones", "Decisiones basadas\nen datos"],
   ["s2_inversion", "Inversión pública\nmás eficiente"],
   ["s2_ciudades", "Ciudades sostenibles\ne inclusivas"]]
  .forEach(([ic, label], i) => {
    const y = 0.85 + i * 0.92;
    s.addShape(pres.ShapeType.ellipse, { x: 9.06, y, w: 0.52, h: 0.52, fill: { color: BLUE }, line: { type: "none" } });
    s.addImage({ path: I(ic), x: 9.19, y: y + 0.13, w: 0.26, h: 0.26 });
    s.addText(label, { x: 9.8, y: y - 0.04, w: 2.95, h: 0.6, fontFace: FONT, fontSize: 11,
      color: INK, margin: 0, valign: "middle", lineSpacing: 14.5 });
  });

  s.addShape(pres.ShapeType.rect, { x: 0, y: 5.12, w: W, h: 2.38, fill: { color: NAVY }, line: { type: "none" } });
  const kw = W / 4;
  [["s2_kpi_municipios", "MUNICIPIOS\nCONECTADOS", "> 50"],
   ["s2_kpi_ciudadanos", "CIUDADANOS\nIMPACTADOS", "> 3M"],
   ["s2_kpi_proyectos", "PROYECTOS\nPRIORIZADOS", "> 100"],
   ["s2_kpi_decision", "MEJOR DECISIÓN\nPÚBLICA", "100%"]]
  .forEach(([ic, label, val], i) => {
    const bx = i * kw;
    s.addImage({ path: I(ic), x: bx + 0.48, y: 5.92, w: 0.66, h: 0.66 });
    s.addText(label, { x: bx + 1.34, y: 5.56, w: kw - 1.55, h: 0.55, fontFace: FONT, fontSize: 10,
      bold: true, color: "FFFFFF", charSpacing: 0.8, margin: 0, valign: "middle", lineSpacing: 13 });
    s.addText(val, { x: bx + 1.34, y: 6.18, w: kw - 1.55, h: 0.62, fontFace: FONT, fontSize: 27,
      bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
    if (i > 0) vline(s, bx, 5.56, 1.5, "FFFFFF", 72);
  });

  s.addNotes("La visión ordena el discurso: conectividad, eficiencia municipal y mejor servicio al ciudadano. Los cuatro indicadores de la banda inferior son las metas de referencia del instrumento.");
}

/* ══════════════════ 3 · LOS ACTORES ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addImage({ path: A("globe_soft.jpg"), x: -0.55, y: 1.34, w: 6.0, h: 5.12 });
  header(s, "Primero: Quiénes son los actores", null, null);

  const X = 6.32, WIDTH = 6.4;
  let y = 1.8;
  [[1, "InDiCo Global", "Proyecto de cooperación internacional sobre estándares digitales/ICT y políticas digitales. Su valor para RD es transferencia de capacidades, alineación con estándares y apoyo técnico."],
   [2, "OASC", "Open & Agile Smart Cities and Communities. Red global de ciudades y comunidades que ayuda a administraciones locales a avanzar en transformación digital e interoperabilidad."],
   [3, "U4SSC", "United for Smart Sustainable Cities. Iniciativa del sistema ONU que ofrece KPIs y modelos de medición para ciudades inteligentes y sostenibles."]]
  .forEach(([n, t, b], i, arr) => {
    const r = numItem(s, { x: X, y, w: WIDTH, num: n, title: t, body: b, gap: 0.5 });
    // Línea que encadena los círculos, como en el original.
    if (i < arr.length - 1) vline(s, X + 0.23, y + 0.46, r.next - y - 0.46, "AFC2DA");
    y = r.next;
  });

  band(s, "Resumen estratégico: InDiCo abre la cooperación; OASC aporta interoperabilidad y red; U4SSC aporta medición, INDOTEL convierte todo en política pública operativa.", 6.12, 0.74, 3);
  sources(s, "Fuentes: InDiCo Global; OASC; U4SSC / ITU.");
  s.addNotes("Tres actores, tres aportes distintos y complementarios. Evita la confusión más común: creer que InDiCo, OASC y U4SSC son lo mismo.");
}

/* ══════════════════ 4 · InDiCo GLOBAL ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  header(s, "InDiCo Global: cooperación digital y estándares",
    "No es solo un “proyecto”; es una puerta de acceso a conocimiento y estándares internacionales.", null);

  const colW = 5.28, x2 = 6.95;
  vline(s, 6.55, 2.1, 3.55, RULE);
  [["Qué promueve", M, [
      "Conocimiento sobre modelos de estándares ICT/digitales.",
      "Capacitación y cooperación con países y regiones socias.",
      "Interoperabilidad, ciberseguridad, datos, IoT, 5G, IA y ciudades inteligentes.",
      "Uso de estándares compatibles con valores europeos: apertura, protección de datos y confianza."]],
   ["Aplicación directa en RD", x2, [
      "OASC trabajará con 3–5 ciudades dominicanas para apoyar su madurez digital.",
      "Se usarán LORDIMAS y el modelo ITU de madurez de ciudades inteligentes y sostenibles.",
      "Meta publicada: crear un hub sostenible vinculado a OASC y U4SSC."]]]
  .forEach(([title, cx, items]) => {
    s.addText(title, { x: cx, y: 2.1, w: colW, h: 0.32, fontFace: FONT, fontSize: 13.5, bold: true,
      color: SUB, margin: 0, valign: "middle" });
    iconList(s, items, { x: cx, y: 2.66, w: colW, icon: "check_blue", fontSize: 11, gap: 0.36 });
  });

  band(s, "Para INDOTEL: esto permite pasar de una cooperación puntual a una arquitectura nacional de acompañamiento municipal.", 6.02, 0.74, 4);
  sources(s, "Fuentes: InDiCo Global; InDiCo-Global Second Open Call Winners, 2026.");
  s.addNotes("InDiCo Global no aporta solo financiamiento: aporta acceso a estándares, capacitación y una red de pares. En RD ya hay metas concretas: 3 a 5 ciudades piloto y un hub sostenible.");
}

/* ══════════════════ 5 · OASC ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  header(s, "OASC: Ciudades Inteligentes Abiertas y Ágiles",
    "Open & Agile Smart Cities and Communities es una red global de ciudades y comunidades.", null);

  const colW = CW / 3;
  [["s5_naturaleza", BLUE, "Naturaleza", "Organización internacional con sede en Bruselas que conecta ciudades, regiones y comunidades para compartir soluciones digitales basadas en datos."],
   ["s5_tecnico", "0A2E7A", "Enfoque técnico", "Trabaja la interoperabilidad técnica, semántica, organizacional y legal, para que sistemas y datos puedan combinarse y reutilizarse."],
   ["s5_valor", "0E7A6B", "Valor público", "Busca evitar dependencia de proveedores, reducir costos, mejorar eficiencia y facilitar la réplica de soluciones entre ciudades."]]
  .forEach(([ic, col, title, body], i) => {
    const cx = M + i * colW, mid = cx + colW / 2;
    s.addShape(pres.ShapeType.ellipse, { x: mid - 0.47, y: 2.24, w: 0.94, h: 0.94, fill: { color: col }, line: { type: "none" } });
    s.addImage({ path: I(ic), x: mid - 0.25, y: 2.46, w: 0.5, h: 0.5 });
    s.addText(title, { x: cx, y: 3.4, w: colW, h: 0.34, fontFace: FONT, fontSize: 14.5, bold: true,
      color: SUB, align: "center", margin: 0, valign: "middle" });
    s.addText(body, { x: cx + 0.42, y: 3.86, w: colW - 0.84, h: 1.4, fontFace: FONT, fontSize: 11,
      color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 16 });
    if (i > 0) vline(s, cx, 2.3, 2.95, RULE);
  });

  s.addShape(pres.ShapeType.roundRect, { x: M + 0.85, y: 5.5, w: CW - 1.7, h: 0.95, rectRadius: 0.07,
    fill: { color: "E4EDF9" }, line: { type: "none" } });
  s.addText("OASC no vende una plataforma cerrada; impulsa reglas comunes\npara que las plataformas municipales puedan comunicarse.", {
    x: M + 1.1, y: 5.5, w: CW - 2.2, h: 0.95, fontFace: FONT, fontSize: 12.5, bold: true,
    color: SUB, align: "center", margin: 0, valign: "middle", lineSpacing: 18 });

  sources(s, "Fuentes: OASC official site; Living-in.eu OASC partner profile.");
  pageNum(s, 5);
  s.addNotes("El punto clave: OASC no es un proveedor, es una red que define reglas comunes. Eso protege al Estado de la dependencia tecnológica.");
}

/* ══════════════════ 6 · CÓMO OPERA OASC ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addImage({ path: A("rooftop_soft.jpg"), x: 6.25, y: 1.76, w: 8.6, h: 5.74 });
  header(s, "Cómo opera OASC",
    "La organización combina gobernanza de ciudades, soporte técnico y participación en proyectos internacionales.",
    6, { subW: 6.0 });

  let y = 2.2;
  [["s6_gobernanza", "Gobernanza", "Consejo de Ciudades, órgano de gobierno y Consejo Tecnológico para validar el rumbo y la adopción de MIMs."],
   ["s6_servicios", "Servicios", "Conocimiento, proyectos, herramientas técnicas, apoyo a miembros, capacitación y espacios de colaboración."],
   ["s6_proyectos", "Proyectos", "Participa en iniciativas de datos, gemelos digitales, CitiVerse, ciberseguridad, movilidad y ciudades inteligentes."],
   ["s6_itu", "Relación con ITU", "En 2024 OASC y la UIT firmaron cooperación para llevar la voz de ciudades y comunidades a tecnologías emergentes."]]
  .forEach(([ic, t, b]) => {
    y = numItem(s, { x: M, y, w: 5.9, icon: ic, square: true, title: t, body: b, gap: 0.42, titleSize: 13 }).next;
  });

  s.addNotes("Cuatro mecanismos de operación. El acuerdo con la UIT de 2024 da respaldo institucional al vínculo con el sistema de Naciones Unidas.");
}

/* ══════════════════ 7 · CONCEPTO CENTRAL: MIMs ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  header(s, "El concepto central de OASC: Mecanismos Mínimos de Interoperabilidad",
    "Minimal Interoperability Mechanisms", null, { titleSize: 19.5, subY: 1.28 });

  const colW = 5.28, x2 = 6.95;
  vline(s, 6.55, 2.0, 3.5, RULE);
  [["¿Qué resuelven?", M, [
      "Permiten un nivel mínimo y suficiente de interoperabilidad entre datos, sistemas y servicios.",
      "Combinan soluciones técnicas y no técnicas.",
      "Ayudan a que las inversiones digitales sean más futuras y menos dependientes de un proveedor.",
      "Mejoran la cooperación entre departamentos y servicios públicos."]],
   ["Interpretación para INDOTEL", x2, [
      "Un municipio no debe comprar sistemas aislados.",
      "Los datos de movilidad, ambiente, permisos, seguridad y conectividad deben poder integrarse.",
      "INDOTEL puede promover lenguaje común, guías técnicas y criterios mínimos para proyectos municipales."]]]
  .forEach(([title, cx, items]) => {
    s.addText(title, { x: cx, y: 2.0, w: colW, h: 0.32, fontFace: FONT, fontSize: 13.5, bold: true,
      color: SUB, margin: 0, valign: "middle" });
    iconList(s, items, { x: cx, y: 2.56, w: colW, icon: "check_blue", fontSize: 11, gap: 0.28 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.72, w: CW, h: 0.95, rectRadius: 0.07,
    fill: { color: "E4EDF9" }, line: { type: "none" } });
  s.addShape(pres.ShapeType.ellipse, { x: M + 0.36, y: 5.97, w: 0.46, h: 0.46, fill: { color: BLUE }, line: { type: "none" } });
  s.addImage({ path: I("quote_shield"), x: M + 0.48, y: 6.09, w: 0.22, h: 0.22 });
  s.addText("“Los MIMs convierten la ciudad inteligente en gestión pública más eficiente,\nno en tecnología aislada.”", {
    x: M + 1.02, y: 5.72, w: CW - 1.4, h: 0.95, fontFace: FONT, fontSize: 12.5, bold: true,
    color: SUB, margin: 0, valign: "middle", lineSpacing: 18 });

  sources(s, "Fuentes: OASC MIMs concept and materials.");
  pageNum(s, 7);
  s.addNotes("Traducción del concepto técnico a decisión de política pública: el mínimo común obligatorio evita compras aisladas y protege la inversión municipal.");
}

/* ══════════════════ 8 · ¿CÓMO OPERA OASC? ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  header(s, "¿Cómo opera OASC?", "Gobernanza, conocimiento y proyectos para ciudades más inteligentes.", 8);

  const colW = CW / 4;
  [["s8_gobernanza", "GOBERNANZA", "Consejo de Ciudades y Consejo Tecnológico para validar el rumbo y estándares (MIMs)."],
   ["s8_servicios", "SERVICIOS", "Conocimiento, herramientas técnicas, capacitación y apoyo continuo a municipios."],
   ["s8_proyectos", "PROYECTOS", "Iniciativas en datos, gemelos digitales, movilidad, ciberseguridad, CitiVerse y más."],
   ["s8_alianzas", "ALIANZAS", "Cooperación con ITU y actores globales para llevar la voz de las ciudades a tecnologías emergentes."]]
  .forEach(([ic, title, body], i) => {
    const cx = M + i * colW;
    s.addImage({ path: I(ic), x: cx + colW / 2 - 0.55, y: 2.55, w: 1.1, h: 1.1 });
    s.addText(title, { x: cx, y: 3.95, w: colW, h: 0.34, fontFace: FONT, fontSize: 13, bold: true,
      color: SUB, align: "center", charSpacing: 0.9, margin: 0, valign: "middle" });
    s.addText(body, { x: cx + 0.32, y: 4.42, w: colW - 0.64, h: 1.4, fontFace: FONT, fontSize: 11.5,
      color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 16 });
    if (i > 0) vline(s, cx, 2.62, 3.2, RULE);
  });

  s.addNotes("Versión sintética de la lámina 6, pensada para audiencias ejecutivas: cuatro pilares, una línea cada uno.");
}

/* ══════════════════ 9 · EL CONCEPTO CLAVE: MIMs ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  s.addImage({ path: A("mims_soft.jpg"), x: -0.5, y: 1.88, w: 7.7, h: 5.14 });
  header(s, "El concepto clave: MIMs",
    "Mecanismos Mínimos de Interoperabilidad. La base para ciudades que se entienden.", 9, { subW: 9.8 });

  iconList(s, [
    "Nivel mínimo y suficiente de interoperabilidad.",
    "Soluciones técnicas y no técnicas.",
    "Inversiones más futuras y menos dependientes.",
    "Lenguaje común, guías y criterios mínimos para todos los municipios."],
    { x: 8.1, y: 2.72, w: 4.6, icon: "check_green", fontSize: 12, gap: 0.55, indent: 0.4 });

  s.addNotes("Lámina de anclaje visual: los MIMs conectan movilidad, ambiente, permisos, seguridad y conectividad en un mismo lenguaje.");
}

/* ══════════════════ 10 · HOJA DE RUTA ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: BG };
  header(s, "Hoja de ruta INDOTEL", "De la cooperación a la transformación nacional.", 10);

  const steps = [
    ["DIAGNÓSTICO", "Evaluamos madurez digital municipal con LORDIMAS e ITU."],
    ["ESTÁNDARES", "Definimos MIMs, guías y arquitectura nacional."],
    ["CAPACIDADES", "Formamos equipos y acompañamos a municipios."],
    ["PROYECTOS", "Desplegamos casos prioritarios con impacto rápido."],
    ["MEDICIÓN", "Medimos resultados con KPIs U4SSC y mejora continua."]];
  const tones = ["0A2456", "0A3070", "0C4088", "1055A8", "2E86D8"];
  const colW = CW / 5, cy = 3.36, d = 0.88;
  const ARROW = { color: "8FA9C6", width: 1.5, endArrowType: "triangle" };

  const cxs = steps.map((_, i) => M + i * colW + colW / 2);
  // Conectores: flecha a cada lado del recorrido y entre cada par de hitos.
  s.addShape(pres.ShapeType.line, { x: M, y: cy, w: cxs[0] - d / 2 - M, h: 0,
    line: { color: "8FA9C6", width: 1.5, beginArrowType: "triangle" } });
  s.addShape(pres.ShapeType.line, { x: cxs[4] + d / 2, y: cy, w: W - M - (cxs[4] + d / 2), h: 0, line: Object.assign({}, ARROW) });
  for (let i = 0; i < 4; i++) {
    s.addShape(pres.ShapeType.line, { x: cxs[i] + d / 2, y: cy, w: colW - d, h: 0, line: Object.assign({}, ARROW) });
  }

  steps.forEach(([title, body], i) => {
    const cx = cxs[i];
    s.addShape(pres.ShapeType.ellipse, { x: cx - d / 2, y: cy - d / 2, w: d, h: d,
      fill: { color: tones[i] }, line: { type: "none" } });
    s.addText(String(i + 1), { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fontFace: FONT, fontSize: 26,
      bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    s.addText(title, { x: cx - colW / 2, y: 4.12, w: colW, h: 0.32, fontFace: FONT, fontSize: 12.5,
      bold: true, color: SUB, align: "center", charSpacing: 0.7, margin: 0, valign: "middle" });
    s.addText(body, { x: cx - colW / 2 + 0.24, y: 4.6, w: colW - 0.48, h: 1.3, fontFace: FONT,
      fontSize: 11.5, color: BODY, align: "center", margin: 0, valign: "top", lineSpacing: 16 });
    if (i > 0) vline(s, cx - colW / 2, 3.98, 2.05, RULE);
  });

  s.addNotes("Cinco fases secuenciales. La medición cierra el ciclo y lo realimenta: sin KPIs U4SSC no hay mejora continua ni evidencia para el Consejo Directivo.");
}

/* ══════════════════ 11 · IMPACTO PARA LA CIUDADANÍA ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: DEEP };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: DEEP }, line: { type: "none" } });
  s.addImage({ path: A("couple_soft.jpg"), x: 7.35, y: 0, w: 6.0, h: 7.48 });
  header(s, null, null, null, { dark: true });

  s.addText("Impacto para la ciudadanía", {
    x: M, y: 1.72, w: 6.5, h: 0.58, fontFace: FONT, fontSize: 26, bold: true,
    color: "FFFFFF", margin: 0, valign: "middle" });
  s.addText("Mejores servicios. Mejor calidad de vida.", {
    x: M, y: 2.42, w: 6.5, h: 0.5, fontFace: FONT, fontSize: 19, bold: true,
    color: "FFFFFF", margin: 0, valign: "middle" });
  // Filete de acento bajo el titular (presente en el original).
  s.addShape(pres.ShapeType.rect, { x: M, y: 3.18, w: 1.5, h: 0.045, fill: { color: SKY }, line: { type: "none" } });

  const iw = 6.55 / 6;
  [["s11_salud", "Salud\ndigital"], ["s11_movilidad", "Movilidad\ninteligente"],
   ["s11_seguridad", "Seguridad\nciudadana"], ["s11_alumbrado", "Alumbrado\neficiente"],
   ["s11_ambiente", "Medio ambiente\nsostenible"], ["s11_digital", "Servicios\ndigitales 24/7"]]
  .forEach(([ic, label], i) => {
    const cx = M + i * iw;
    s.addImage({ path: I(ic), x: cx + iw / 2 - 0.24, y: 3.86, w: 0.48, h: 0.48 });
    s.addText(label, { x: cx + 0.03, y: 4.5, w: iw - 0.06, h: 0.66, fontFace: FONT, fontSize: 9.5,
      color: "DCE7F5", align: "center", margin: 0, valign: "top", lineSpacing: 12.5 });
    if (i > 0) vline(s, cx, 3.78, 1.5, "FFFFFF", 80);
  });

  pageNum(s, 11, true);
  s.addNotes("Cierre del argumento técnico con el argumento humano: seis servicios cotidianos donde el ciudadano percibe el resultado de la interoperabilidad.");
}

/* ══════════════════ 12 · DECISIÓN ESTRATÉGICA ══════════════════ */
{
  const s = pres.addSlide();
  s.background = { color: DEEP };
  s.addImage({ path: A("city_dark.jpg"), x: 0, y: 2.35, w: W, h: 5.15, sizing: { type: "cover", w: W, h: 5.15 } });
  s.addImage({ path: A("veil_close.png"), x: 0, y: 0, w: W, h: H });
  header(s, null, null, null, { dark: true });

  s.addText("Decisión estratégica", {
    x: M, y: 1.32, w: 6, h: 0.36, fontFace: FONT, fontSize: 15, color: "D6E4F5", margin: 0, valign: "middle" });
  s.addText("Autorización para convertir esta cooperación en un instrumento permanente del INDOTEL.", {
    x: M, y: 1.8, w: 8.35, h: 1.8, fontFace: FONT, fontSize: 28, bold: true,
    color: "FFFFFF", margin: 0, valign: "top", lineSpacing: 36 });
  s.addText("Más ciudades conectadas. Más datos que generan valor.\nMejores decisiones. Un país que avanza.", {
    x: M, y: 3.78, w: 7.8, h: 0.95, fontFace: FONT, fontSize: 15,
    color: "DCE7F5", margin: 0, valign: "top", lineSpacing: 24 });

  vline(s, 9.32, 1.3, 4.35, "FFFFFF", 66);

  s.addText("INDOTEL", { x: 9.85, y: 1.72, w: 3.1, h: 0.55, fontFace: MARK, fontSize: 29, bold: true,
    color: "FFFFFF", charSpacing: 1.8, margin: 0, valign: "middle" });
  s.addText("Lideramos la transformación digital de los municipios dominicanos.", {
    x: 9.85, y: 2.44, w: 2.95, h: 1.1, fontFace: FONT, fontSize: 12.5,
    color: "DCE7F5", margin: 0, valign: "top", lineSpacing: 18.5 });
  s.addText("2026", { x: 9.85, y: 3.88, w: 3.1, h: 0.62, fontFace: FONT, fontSize: 30, bold: true,
    color: "FFFFFF", margin: 0, valign: "middle" });
  s.addShape(pres.ShapeType.rect, { x: 9.85, y: 4.56, w: 1.05, h: 0.05, fill: { color: SKY }, line: { type: "none" } });

  pageNum(s, 12, true);
  s.addNotes("Lámina de cierre y solicitud formal: el Consejo Directivo autoriza convertir la cooperación InDiCo/OASC/U4SSC en un instrumento permanente del INDOTEL.");
}

pres.writeFile({ fileName: path.join(__dirname, "..", "INDOTEL-Estrategia-Nacional-Ciudades-Inteligentes.pptx") })
  .then((f) => console.log("OK ->", f));
