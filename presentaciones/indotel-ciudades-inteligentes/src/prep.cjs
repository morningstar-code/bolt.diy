const sharp = require("sharp");
const path = require("path");
const fs = require("fs");

const SRC = (f) => path.join(__dirname, "assets", f);
const OUT = (f) => path.join(__dirname, "derived", f);
fs.mkdirSync(path.join(__dirname, "derived"), { recursive: true });

// Fondos sobre los que se aplanan las fotos (deben coincidir con build.cjs)
const BG   = { r: 242, g: 243, b: 248 }; // láminas claras
const DEEP = { r: 4,   g: 23,  b: 63  }; // láminas oscuras

/* ── Fundidos ─────────────────────────────────────────────────────────────
   El original no usa fotos con borde recto: se desvanecen hacia el fondo.
   Se genera una máscara de alfa por lado y se aplica con `dest-in`.        */
function mask(w, h, side, frac) {
  const g = {
    left:   `x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-opacity="0"/><stop offset="${frac}" stop-opacity="1"/>`,
    right:  `x1="0" y1="0" x2="1" y2="0"><stop offset="${1 - frac}" stop-opacity="1"/><stop offset="1" stop-opacity="0"/>`,
    top:    `x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-opacity="0"/><stop offset="${frac}" stop-opacity="1"/>`,
    bottom: `x1="0" y1="0" x2="0" y2="1"><stop offset="${1 - frac}" stop-opacity="1"/><stop offset="1" stop-opacity="0"/>`,
  }[side];
  return Buffer.from(
    `<svg width="${w}" height="${h}"><defs><linearGradient id="g" ${g}</linearGradient></defs>` +
    `<rect width="100%" height="100%" style="fill:url(#g)"/></svg>`
  );
}

async function feather(buf, sides) {
  let out = await sharp(buf).ensureAlpha().png().toBuffer();
  const { width: w, height: h } = await sharp(out).metadata();
  for (const [side, frac] of sides) {
    out = await sharp(out)
      .composite([{ input: mask(w, h, side, frac), blend: "dest-in" }])
      .png().toBuffer();
  }
  return out;
}

// Aplanar sobre el color de fondo de la lámina: mismo resultado visual que
// el alfa, sin canal alfa -> JPEG, y el archivo pesa una fracción.
async function flattenOn(buf, bg, outName, width) {
  // Redimensionar primero: sharp aplica `resize` antes que `composite`, así que
  // el lienzo y la capa deben llegar ya del mismo tamaño.
  const small = await sharp(buf).resize({ width, withoutEnlargement: true }).png().toBuffer();
  const { width: w, height: h } = await sharp(small).metadata();
  await sharp({ create: { width: w, height: h, channels: 3, background: bg } })
    .composite([{ input: small }])
    .jpeg({ quality: 86, mozjpeg: true })
    .toFile(OUT(outName));
}

(async () => {
  /* ── Lámina 2: monumento fundido hacia la izquierda ───────────────────── */
  await flattenOn(
    await feather(await sharp(SRC("monumento.jpg")).toBuffer(), [["left", 0.20]]),
    BG, "monumento_soft.jpg", 1200
  );

  /* ── Lámina 3: globo recortado y fundido por tres lados ───────────────── */
  {
    const m = await sharp(SRC("globe.jpg")).metadata();
    const crop = await sharp(SRC("globe.jpg"))
      .extract({ left: 0, top: 0, width: Math.round(m.width * 0.66), height: m.height })
      .toBuffer();
    await flattenOn(
      await feather(crop, [["right", 0.40], ["top", 0.30], ["bottom", 0.30]]),
      BG, "globe_soft.jpg", 1120
    );
  }

  /* ── Lámina 6: azotea, más luminosa y fundida hacia el texto ──────────── */
  {
    const light = await sharp(SRC("rooftop.jpg"))
      .modulate({ brightness: 1.14, saturation: 0.88 })
      .toBuffer();
    await flattenOn(
      await feather(light, [["left", 0.30], ["top", 0.32]]),
      BG, "rooftop_soft.jpg", 1120
    );
  }

  /* ── Lámina 9: diagrama MIMs; sangra por izquierda y abajo, corte recto
       a la derecha (como el original). Solo se funde el borde superior.  ── */
  await flattenOn(
    await feather(await sharp(SRC("mims.jpg")).toBuffer(), [["top", 0.18], ["bottom", 0.14]]),
    BG, "mims_soft.jpg", 1220
  );

  /* ── Lámina 11: la pareja. El asset trae el chip "AMBIENTE" superpuesto
       sobre el brazo del hombre; se difumina bajo una máscara de bordes
       suaves antes de recortar.                                          ── */
  {
    const { width: mw, height: mh } = await sharp(SRC("mims.jpg")).metadata();
    const px = (f, d) => Math.round(f * d);
    const patch = { left: px(0.424, mw), top: px(0.412, mh), width: px(0.152, mw), height: px(0.180, mh) };
    const soft = Buffer.from(
      `<svg width="${patch.width}" height="${patch.height}">` +
      `<defs><filter id="f" x="-30%" y="-30%" width="160%" height="160%">` +
      `<feGaussianBlur stdDeviation="${Math.round(patch.width * 0.11)}"/></filter></defs>` +
      `<rect x="${patch.width * 0.14}" y="${patch.height * 0.16}" ` +
      `width="${patch.width * 1.1}" height="${patch.height * 0.68}" ` +
      `rx="${patch.width * 0.12}" fill="#fff" filter="url(#f)"/></svg>`
    );
    const blurred = await sharp(SRC("mims.jpg")).extract(patch).blur(34).ensureAlpha()
      .composite([{ input: soft, blend: "dest-in" }]).png().toBuffer();
    const patched = await sharp(SRC("mims.jpg"))
      .composite([{ input: blurred, left: patch.left, top: patch.top }])
      .png().toBuffer();
    const cropped = await sharp(patched)
      .extract({ left: 0, top: 0, width: px(0.535, mw), height: mh })
      .png().toBuffer();
    await flattenOn(
      await feather(cropped, [["right", 0.14], ["left", 0.13], ["bottom", 0.11], ["top", 0.09]]),
      DEEP, "couple_soft.jpg", 940
    );
  }

  /* ── Portada y cierre ─────────────────────────────────────────────────── */
  await sharp(SRC("city_night.jpg")).resize({ width: 2000, withoutEnlargement: true })
    .jpeg({ quality: 85, mozjpeg: true }).toFile(OUT("city_night.jpg"));
  await sharp(SRC("city_night.jpg")).modulate({ brightness: 0.52 })
    .resize({ width: 2000, withoutEnlargement: true })
    .jpeg({ quality: 85, mozjpeg: true }).toFile(OUT("city_dark.jpg"));

  // Velo degradado: un rectángulo semitransparente dejaría una costura vertical.
  const veil = (stops) => Buffer.from(
    `<svg width="1600" height="900"><defs><linearGradient id="v" x1="0" y1="0" x2="1" y2="0">${stops}</linearGradient>` +
    `</defs><rect width="100%" height="100%" fill="url(#v)"/></svg>`
  );
  await sharp(veil(
    `<stop offset="0" stop-color="#04173F" stop-opacity="0.93"/>` +
    `<stop offset="0.44" stop-color="#04173F" stop-opacity="0.74"/>` +
    `<stop offset="0.80" stop-color="#04173F" stop-opacity="0.34"/>` +
    `<stop offset="1" stop-color="#04173F" stop-opacity="0.22"/>`
  )).png().toFile(OUT("veil_cover.png"));
  await sharp(veil(
    `<stop offset="0" stop-color="#04173F" stop-opacity="0.95"/>` +
    `<stop offset="0.55" stop-color="#04173F" stop-opacity="0.80"/>` +
    `<stop offset="1" stop-color="#04173F" stop-opacity="0.58"/>`
  )).png().toFile(OUT("veil_close.png"));

  console.log("Assets preparados.");
})();
