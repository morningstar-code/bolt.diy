const sharp = require("sharp");
const path = require("path");

const SRC = (f) => path.join(__dirname, "assets", f);
const A = (f) => path.join(__dirname, "derived", f);
require("fs").mkdirSync(path.join(__dirname, "derived"), { recursive: true });

// Máscara de desvanecido lateral: alpha 0 en el borde indicado -> 1 hacia adentro.
function fadeMask(w, h, sides, frac = 0.3) {
  const stops = {
    left:  `<linearGradient id="g" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-opacity="0"/><stop offset="${frac}" stop-opacity="1"/></linearGradient>`,
    right: `<linearGradient id="g" x1="0" y1="0" x2="1" y2="0"><stop offset="${1 - frac}" stop-opacity="1"/><stop offset="1" stop-opacity="0"/></linearGradient>`,
    bottom:`<linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="${1 - frac}" stop-opacity="1"/><stop offset="1" stop-opacity="0"/></linearGradient>`,
    top:   `<linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-opacity="0"/><stop offset="${frac}" stop-opacity="1"/></linearGradient>`,
  };
  return Buffer.from(
    `<svg width="${w}" height="${h}"><defs>${stops[sides]}</defs>` +
    `<rect width="100%" height="100%" fill="#000" fill-opacity="1" mask="none" style="fill:url(#g)"/></svg>`
  );
}

(async () => {
  // --- Slide 11: recorte de la pareja. El asset trae el chip "AMBIENTE"
  // superpuesto sobre el brazo del hombre; se difumina antes de recortar.
  const { width: mw, height: mh } = await sharp(SRC("mims.jpg")).metadata();
  const px = (f, d) => Math.round(f * d);

  const patch = { left: px(0.424, mw), top: px(0.412, mh), width: px(0.152, mw), height: px(0.180, mh) };
  // Máscara con bordes difuminados para que el parche funda con el original.
  const feather = Buffer.from(
    `<svg width="${patch.width}" height="${patch.height}">` +
    `<defs><filter id="f" x="-30%" y="-30%" width="160%" height="160%">` +
    `<feGaussianBlur stdDeviation="${Math.round(patch.width * 0.11)}"/></filter></defs>` +
    `<rect x="${patch.width * 0.14}" y="${patch.height * 0.16}" ` +
    `width="${patch.width * 1.1}" height="${patch.height * 0.68}" ` +
    `rx="${patch.width * 0.12}" fill="#fff" filter="url(#f)"/></svg>`
  );
  const blurred = await sharp(SRC("mims.jpg")).extract(patch).blur(34).ensureAlpha()
    .composite([{ input: feather, blend: "dest-in" }]).png().toBuffer();
  let couple = await sharp(SRC("mims.jpg"))
    .composite([{ input: blurred, left: patch.left, top: patch.top }])
    .png().toBuffer();

  const cw = px(0.535, mw);
  couple = await sharp(couple).extract({ left: 0, top: 0, width: cw, height: mh }).ensureAlpha().png().toBuffer();
  for (const [side, frac] of [["right", 0.13], ["left", 0.13], ["bottom", 0.10], ["top", 0.09]]) {
    couple = await sharp(couple)
      .composite([{ input: fadeMask(cw, mh, side, frac), blend: "dest-in" }])
      .png().toBuffer();
  }
  // Lienzo transparente a los lados: así el encuadre de PowerPoint nunca recorta
  // los fundidos y la foto no muestra bordes duros sobre el navy.
  await sharp(couple)
    .extend({ left: 140, right: 139, background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(A("couple_fade.png"));

  // --- Velo degradado navy -> transparente para portada y cierre (evita costuras)
  const gw = 1600, gh = 900;
  const veil = Buffer.from(
    `<svg width="${gw}" height="${gh}"><defs><linearGradient id="v" x1="0" y1="0" x2="1" y2="0">` +
    `<stop offset="0" stop-color="#04162E" stop-opacity="0.92"/>` +
    `<stop offset="0.42" stop-color="#04162E" stop-opacity="0.72"/>` +
    `<stop offset="0.78" stop-color="#04162E" stop-opacity="0.34"/>` +
    `<stop offset="1" stop-color="#04162E" stop-opacity="0.24"/>` +
    `</linearGradient></defs><rect width="100%" height="100%" fill="url(#v)"/></svg>`
  );
  await sharp(veil).png().toFile(A("veil_left.png"));

  // --- Slide 3: globo con fundido a la derecha (el original ya aclara hacia el borde)
  const g = await sharp(SRC("globe.jpg")).metadata();
  await sharp(SRC("globe.jpg"))
    .extract({ left: 0, top: 0, width: Math.round(g.width * 0.62), height: g.height })
    .png()
    .toFile(A("globe_crop.png"));

  // --- Slide 12: skyline nocturno oscurecido para fondo de cierre
  await sharp(SRC("city_night.jpg"))
    .modulate({ brightness: 0.55 })
    .png()
    .toFile(A("city_dark.png"));

  // --- Compresión: el deck debe poder enviarse por correo.
  // Ancho objetivo ~150 DPI del tamaño real en la lámina.
  const OPT = [
    ["__src_city_night", "city_night.jpg", 2000],
    ["city_dark.png",  "city_dark.jpg",  2000],
    ["__src_monumento",  "monumento.jpg",  1150],
    ["globe_crop.png", "globe_crop.jpg",  980],
    ["__src_rooftop",    "rooftop.jpg",     920],
    ["__src_mims",       "mims.jpg",       1250],
  ];
  for (const [src, dst, w] of OPT) {
    await sharp(src.startsWith("__src_") ? SRC(src.slice(6) + ".jpg") : A(src)).resize({ width: w, withoutEnlargement: true })
      .jpeg({ quality: 84, mozjpeg: true }).toFile(A(dst));
  }
  // Conserva alfa: PNG, pero redimensionado.
  await sharp(A("couple_fade.png")).resize({ width: 880 })
    .png({ compressionLevel: 9, palette: false }).toFile(A("couple_opt.png"));

  console.log("Assets preparados y optimizados.");
})();
