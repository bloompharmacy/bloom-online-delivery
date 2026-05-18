/**
 * Bloom Online Delivery — Firestore Matrix Seed Script
 * ─────────────────────────────────────────────────────
 * Writes the complete 47×47 symmetric shipping-fee matrix to
 *   /shipping_fees/matrix
 * in the dedicated isolated Firestore database.
 *
 * HOW TO RUN (one-time admin operation):
 *   1. Open ol-admin.html → Settings tab → "Seed / Reset Matrix"
 *   2. OR paste into a browser console after initialising Firebase
 *      and call: seedShippingMatrix('<your-admin-passcode>')
 *
 * Data source: Shippingfees.xlsx (47 regions × 47 regions)
 * All values are in EGP.
 * Diagonal (same → same region) = minimum local delivery fee, NOT zero.
 * Matrix is symmetric: cost[A][B] === cost[B][A] ✓
 */

// ─── Region registry (order determines dropdown sort in audit UI) ─────────────

const REGIONS = [
  { id: 'mohandseen',         label: 'Mohandseen',                    order:  0 },
  { id: 'dokki',              label: 'Dokki',                         order:  1 },
  { id: 'zamalek',            label: 'Zamalek',                       order:  2 },
  { id: 'bulaq_daqrur',       label: 'Bulaq Daqrur',                  order:  3 },
  { id: 'giza',               label: 'Giza',                          order:  4 },
  { id: 'imbaba',             label: 'Imbaba',                        order:  5 },
  { id: 'garden_city',        label: 'Garden City',                   order:  6 },
  { id: 'hadayek_el_kobba',   label: 'Hadayek el Kobba',              order:  7 },
  { id: 'al_haram',           label: 'Al Haram',                      order:  8 },
  { id: 'shobra_al_khayma',   label: 'Shobra al Khayma',              order:  9 },
  { id: 'el_zaytoun',         label: 'El Zaytoun',                    order: 10 },
  { id: 'maadi',              label: 'Maadi',                         order: 11 },
  { id: 'heliopolis',         label: 'Heliopolis',                    order: 12 },
  { id: 'mokattam',           label: 'Mokattam',                      order: 13 },
  { id: 'nasr_city',          label: 'Nasr City',                     order: 14 },
  { id: 'ain_shams',          label: 'Ain Shams',                     order: 15 },
  { id: 'al_marj',            label: 'Al Marj',                       order: 16 },
  { id: 'al_hadaba_el_wosta', label: 'Al Hadaba el Wosta',            order: 17 },
  { id: 'al_harfyeen',        label: 'Al Harfyeen',                   order: 18 },
  { id: 'sheikh_zayed',       label: 'Sheikh Zayed',                  order: 19 },
  { id: 'el_salam_city',      label: 'El Salam City',                 order: 20 },
  { id: 'helwan',             label: 'Helwan',                        order: 21 },
  { id: 'al_rehab',           label: 'Al Rehab / 1st Settlement',     order: 22 },
  { id: 'new_cairo_3rd_5th',  label: 'New Cairo (3rd & 5th Settlement)', order: 23 },
  { id: 'sixth_october',      label: '6th of October',                order: 24 },
  { id: 'el_obour_city',      label: 'El Obour City',                 order: 25 },
  { id: 'madinaty',           label: 'Madinaty',                      order: 26 },
  { id: 'el_shorouk',         label: 'El Shorouk',                    order: 27 },
  { id: 'badr',               label: 'Badr City',                     order: 28 },
  { id: 'tenth_ramadan',      label: '10th of Ramadan',               order: 29 },
  { id: 'capital_city',       label: 'Capital City',                  order: 30 },
  { id: 'al_mansouria',       label: 'Al Mansouria',                  order: 31 },
  { id: 'misr_el_qadima',     label: 'Misr El Qadima',                order: 32 },
  { id: 'shobra',             label: 'Shobra',                        order: 33 },
  { id: 'sheraton_almaza',    label: 'Sheraton / Almaza',             order: 34 },
  { id: 'airport',            label: 'Airport',                       order: 35 },
  { id: 'manyal',             label: 'Manyal',                        order: 36 },
  { id: 'al_matariyyah',      label: 'Al Matariyyah',                 order: 37 },
  { id: 'new_cairo_auc',      label: 'New Cairo (AUC / Mivida)',       order: 38 },
  { id: 'beverly_hills',      label: 'Beverly Hills',                 order: 39 },
  { id: 'pyramids_gardens',   label: 'Pyramids Gardens',              order: 40 },
  { id: 'october_gardens',    label: 'October Gardens',               order: 41 },
  { id: 'new_giza',           label: 'New Giza',                      order: 42 },
  { id: 'mv_icity',           label: 'New Cairo (Mountain View / iCity)', order: 43 },
  { id: 'zahraa_el_maadi',    label: 'Zahraa el Maadi',               order: 44 },
  { id: 'carfour_el_maadi',   label: 'Carrefour El Maadi',            order: 45 },
  { id: 'zahraa_nasr_city',   label: 'Zahraa Nasr City',              order: 46 },
];

// ─── Raw matrix rows (source: Shippingfees.xlsx) ─────────────────────────────
// Each row corresponds to REGIONS[rowIndex]. Values are EGP costs to each
// destination in column order (same order as REGIONS array above).

const RAW_MATRIX = [
  // mohandseen
  [ 12,  33,  33,  50,  55,  55,  55,  77,  99, 105, 110, 116, 176, 138, 193, 193, 160, 176, 176, 242, 259, 259, 286, 314, 314, 330, 341, 369, 424, 506, 506, 215,  66,  72, 149, 226,  61, 132, 314, 209, 143, 242, 149, 303, 154, 143, 165],
  // dokki
  [ 33,  12,  44,  44,  66,  66,  44,  77,  83,  83,  99, 105, 132,  99, 154, 149, 176, 138, 187, 193, 198, 215, 253, 281, 281, 297, 308, 336, 391, 473, 473, 209,  66,  72, 143, 176,  50, 116, 281, 226, 116, 231, 171, 297, 154, 138, 160],
  // zamalek
  [ 33,  44,  12,  50,  66,  50,  39,  72, 105,  77,  99,  99, 121,  94, 143, 143, 154, 138, 226, 193, 209, 198, 253, 281, 281, 297, 308, 336, 391, 473, 473, 176,  61,  50, 132, 165,  55, 116, 264, 220, 154, 253, 165, 286, 143, 138, 154],
  // bulaq_daqrur
  [ 50,  44,  50,  12,  72,  55,  61,  94,  99,  83, 110, 105, 127, 110, 182, 160, 182, 143, 209, 198, 215, 220, 253, 281, 281, 297, 308, 336, 391, 473, 473, 171,  66,  77, 165, 198,  66, 132, 297, 237, 116, 220, 154, 308, 154, 149, 171],
  // giza
  [ 55,  66,  66,  72,  12,  83,  61,  99,  94,  99, 121,  66, 143,  88, 171, 171, 193, 110, 198, 231, 226, 171, 253, 253, 281, 308, 336, 363, 418, 473, 473, 215,  83, 105, 154, 187,  44, 143, 292, 259, 138, 198, 182, 303, 121, 116, 165],
  // imbaba
  [ 55,  66,  50,  55,  83,  12,  72,  83, 121,  77, 105, 127, 143, 116, 149, 127, 138, 160, 160, 193, 198, 231, 242, 281, 281, 286, 308, 363, 446, 418, 501, 138,  83,  55, 154, 187,  77, 132, 281, 215, 143, 248, 154, 319, 171, 165, 198],
  // garden_city
  [ 55,  44,  39,  61,  61,  72,  12,  61,  99,  72,  88,  88, 116,  83, 116, 138, 165, 121, 176, 209, 204, 209, 209, 237, 286, 264, 292, 325, 369, 418, 473, 198,  44,  50, 121, 154,  44, 105, 264, 237, 165, 253, 182, 281, 138, 127, 143],
  // hadayek_el_kobba
  [ 77,  77,  72,  94,  99,  83,  61,  25, 132,  72,  50, 121,  72,  88, 110, 105,  94, 138, 143, 237, 165, 226, 198, 220, 319, 242, 275, 308, 347, 391, 418, 198,  55,  55,  94, 127,  83,  66, 231, 270, 198, 303, 209, 242, 160, 143, 121],
  // al_haram
  [ 99,  83, 105,  99,  94, 121,  99, 132,  25, 165, 209, 138, 215, 165, 220, 231, 231, 176, 270, 165, 270, 253, 297, 330, 226, 374, 374, 374, 501, 556, 556, 143, 143, 143, 193, 226,  83, 209, 363, 237,  83, 171, 138, 325, 143, 143, 198],
  // shobra_al_khayma
  [105,  83,  77,  83,  99,  77,  72,  72, 165,  25,  66, 187,  94, 143, 121,  99,  99, 182, 138, 253, 171, 264, 209, 220, 363, 253, 297, 319, 363, 418, 501, 171,  94,  61, 121, 154,  88,  94, 275, 253, 198, 297, 209, 286, 176, 171, 165],
  // el_zaytoun
  [110,  99,  99, 110, 121, 105,  88,  50, 209,  66,  25, 165,  61, 116,  88,  61,  66, 160,  99, 264, 132, 264, 198, 231, 352, 198, 253, 275, 336, 363, 418, 204,  94,  77,  66,  99, 116,  39, 226, 286, 198, 330, 237, 231, 182, 165, 121],
  // maadi
  [116, 105,  99, 105,  66, 127,  88, 121, 138, 187, 165,  25, 171, 105, 176, 209, 226,  94, 231, 275, 253, 143, 253, 231, 308, 336, 336, 363, 418, 501, 446, 248,  99, 132, 182, 253,  72, 193, 253, 297, 165, 226, 209, 319,  66,  88, 182],
  // heliopolis
  [176, 132, 121, 127, 143, 143, 116,  72, 215,  94,  61, 171,  25, 105,  72,  77,  77, 143,  99, 270, 132, 248, 165, 220, 358, 220, 242, 259, 336, 391, 446, 209,  72,  77,  61,  99, 110,  55, 226, 303, 215, 292, 226, 220, 154, 143, 121],
  // mokattam
  [138,  99,  94, 110,  88, 116,  83,  88, 165, 143, 116, 105, 105,  25, 110, 160, 209,  66, 182, 281, 204, 204, 220, 204, 336, 281, 308, 336, 391, 446, 429, 264, 105, 121, 154, 187, 105, 165, 281, 325, 226, 286, 253, 281, 138, 116, 132],
  // nasr_city
  [193, 154, 143, 182, 171, 149, 116, 110, 220, 121,  88, 176,  72, 110,  25, 105, 127, 138, 127, 270, 171, 275, 154, 176, 418, 231, 226, 253, 330, 402, 402, 253, 110, 110,  88, 105, 143, 105, 198, 319, 264, 303, 264, 215, 143, 116,  66],
  // ain_shams
  [193, 149, 143, 160, 171, 127, 138, 105, 231,  99,  61, 209,  77, 160, 105,  25,  50, 171,  88, 286, 116, 275, 198, 248, 391, 198, 281, 264, 363, 374, 473, 286,  99,  99,  72, 105, 127,  33, 253, 308, 292, 336, 248, 253, 193, 182, 143],
  // al_marj
  [160, 176, 154, 182, 193, 138, 165,  94, 231,  99,  66, 226,  77, 209, 127,  50,  25, 198,  88, 325, 105, 308, 226, 281, 407, 198, 275, 264, 363, 374, 473, 242, 121, 110,  99, 132, 143,  44, 281, 314, 275, 363, 286, 264, 215, 253, 193],
  // al_hadaba_el_wosta
  [176, 138, 138, 143, 110, 160, 121, 138, 176, 182, 160,  94, 143,  66, 138, 171, 198,  25, 187, 297, 215, 198, 209, 198, 336, 281, 292, 308, 391, 446, 418, 330,  99, 138, 138, 154,  94, 143, 226, 325, 198, 253, 237, 270, 105,  44, 116],
  // al_harfyeen
  [176, 187, 226, 209, 198, 160, 176, 143, 270, 138,  99, 231,  99, 182, 127,  88,  88, 187,  25, 352,  66, 325, 198, 242, 446, 171, 253, 226, 347, 341, 446, 259, 149, 138,  77, 116, 182,  83, 253, 358, 292, 352, 308, 231, 220, 198, 143],
  // sheikh_zayed
  [242, 193, 193, 198, 231, 193, 209, 237, 165, 253, 264, 275, 270, 281, 270, 286, 325, 297, 352,  25, 374, 352, 391, 418, 182, 473, 473, 501, 567, 638, 693, 226, 226, 237, 292, 336, 237, 292, 446,  61, 182, 182, 110, 446, 303, 275, 380],
  // el_salam_city
  [259, 198, 209, 215, 226, 198, 204, 165, 270, 171, 132, 253, 132, 204, 171, 116, 105, 215,  66, 374,  25, 385, 193, 242, 440, 171, 248, 226, 308, 336, 446, 275, 165, 149,  99, 138, 198,  99, 253, 358, 281, 407, 292, 215, 248, 253, 154],
  // helwan
  [259, 215, 198, 220, 171, 231, 209, 226, 253, 264, 264, 143, 248, 204, 275, 275, 308, 198, 325, 352, 385,  25, 292, 264, 429, 418, 391, 429, 473, 550, 446, 363, 209, 237, 264, 308, 209, 292, 336, 385, 253, 308, 286, 363, 149, 138, 275],
  // al_rehab
  [286, 253, 253, 253, 253, 242, 209, 198, 297, 209, 198, 253, 165, 220, 154, 198, 226, 209, 198, 391, 193, 292,  25, 143, 484, 198, 193, 209, 253, 347, 308, 336, 193, 209, 127, 160, 237, 154, 132, 391, 336, 391, 369, 171, 198, 182,  99],
  // new_cairo_3rd_5th
  [314, 281, 281, 281, 253, 281, 237, 220, 330, 220, 231, 231, 220, 204, 176, 248, 281, 198, 242, 418, 242, 264, 143,  25, 484, 253, 226, 226, 281, 473, 308, 418, 171, 182, 154, 198, 182, 171, 143, 440, 314, 369, 358, 176, 171, 149, 149],
  // sixth_october
  [314, 281, 281, 281, 281, 281, 286, 319, 226, 363, 352, 308, 358, 336, 418, 391, 407, 336, 446, 182, 440, 429, 484, 484,  25, 528, 539, 556, 693, 693, 748, 281, 292, 292, 374, 418, 292, 374, 501, 198, 182, 121, 165, 616, 385, 319, 418],
  // el_obour_city
  [330, 297, 297, 297, 308, 286, 264, 242, 374, 253, 198, 336, 220, 281, 231, 198, 198, 281, 171, 473, 171, 418, 198, 253, 528,  25, 226, 209, 308, 308, 446, 374, 237, 237, 182, 215, 275, 209, 308, 440, 391, 462, 413, 237, 286, 292, 193],
  // madinaty
  [341, 308, 308, 308, 336, 308, 292, 275, 374, 297, 253, 336, 242, 308, 226, 281, 275, 292, 253, 473, 248, 391, 193, 226, 539, 226,  25, 116, 187, 281, 281, 407, 292, 308, 209, 253, 308, 253, 226, 473, 413, 468, 413, 110, 292, 292, 193],
  // el_shorouk
  [369, 336, 336, 336, 363, 363, 325, 308, 374, 319, 275, 363, 259, 336, 253, 264, 264, 308, 226, 501, 226, 429, 209, 226, 556, 209, 116,  25, 132, 226, 308, 429, 319, 319, 237, 281, 347, 264, 270, 495, 440, 495, 440, 127, 308, 303, 215],
  // badr
  [424, 391, 391, 391, 418, 446, 369, 347, 501, 363, 336, 418, 336, 391, 330, 363, 363, 391, 347, 567, 308, 473, 253, 281, 693, 308, 187, 132,  25, 215, 204, 517, 347, 363, 292, 308, 418, 319, 308, 550, 495, 550, 495, 187, 363, 369, 275],
  // tenth_ramadan
  [506, 473, 473, 473, 473, 418, 418, 391, 556, 418, 363, 501, 391, 446, 402, 374, 374, 446, 341, 638, 336, 550, 347, 473, 693, 308, 281, 226, 215,  25, 391, 556, 374, 402, 319, 391, 457, 319, 336, 660, 550, 660, 550, 275, 396, 468, 358],
  // capital_city
  [506, 473, 473, 473, 473, 501, 473, 418, 556, 501, 418, 446, 446, 429, 402, 473, 473, 418, 446, 693, 446, 446, 308, 308, 748, 446, 281, 308, 204, 391,  25, 600, 374, 402, 347, 363, 457, 374, 226, 770, 660, 715, 660, 275, 440, 385, 385],
  // al_mansouria
  [215, 209, 176, 171, 215, 138, 198, 198, 143, 171, 204, 248, 209, 264, 253, 286, 242, 330, 259, 226, 275, 363, 336, 418, 281, 374, 407, 429, 517, 556, 600,  25, 292, 292, 292, 336, 292, 308, 556, 198, 171, 264, 171, 440, 330, 270, 319],
  // misr_el_qadima
  [ 66,  66,  61,  66,  83,  83,  44,  55, 143,  94,  94,  99,  72, 105, 110,  99, 121,  99, 149, 226, 165, 209, 193, 171, 292, 237, 292, 319, 347, 374, 374, 292,  25,  50, 105, 143,  61, 105, 253, 275, 154, 209, 198, 292, 116,  88, 154],
  // shobra
  [ 72,  72,  50,  77, 105,  55,  50,  55, 143,  61,  77, 132,  77, 121, 110,  99, 110, 138, 138, 237, 149, 237, 209, 182, 292, 237, 308, 319, 363, 402, 402, 292,  50,  25, 110, 143,  66,  88, 264, 253, 182, 303, 193, 275, 176, 138, 165],
  // sheraton_almaza
  [149, 143, 132, 165, 154, 154, 121,  94, 193, 121,  66, 182,  61, 154,  88,  72,  99, 138,  77, 292,  99, 264, 127, 154, 374, 182, 209, 237, 292, 319, 347, 292, 105, 110,  25,  77, 138,  77, 237, 325, 275, 319, 264, 209, 187, 171,  94],
  // airport
  [226, 176, 165, 198, 187, 187, 154, 127, 226, 154,  99, 253,  99, 187, 105, 105, 132, 154, 116, 336, 138, 308, 160, 198, 418, 215, 253, 281, 308, 391, 363, 336, 143, 143,  77,  25, 171, 110, 242, 385, 330, 358, 303, 248, 220, 220, 127],
  // manyal
  [ 61,  50,  55,  66,  44,  77,  44,  83,  83,  88, 116,  72, 110, 105, 143, 127, 143,  94, 182, 237, 198, 209, 237, 182, 292, 275, 308, 347, 418, 457, 457, 292,  61,  66, 138, 171,  25, 127, 281, 259, 116, 242, 182, 303, 127, 110, 165],
  // al_matariyyah
  [132, 116, 116, 132, 143, 132, 105,  66, 209,  94,  39, 193,  55, 165, 105,  33,  44, 143,  83, 292,  99, 292, 154, 171, 374, 209, 253, 264, 319, 319, 374, 308, 105,  88,  77, 110, 127,  25, 253, 286, 237, 330, 248, 248, 193, 176, 132],
  // new_cairo_auc
  [314, 281, 264, 297, 292, 281, 264, 231, 363, 275, 226, 253, 226, 281, 198, 253, 281, 226, 253, 446, 253, 336, 132, 143, 501, 308, 226, 270, 308, 336, 226, 556, 253, 264, 237, 242, 281, 253,  25, 495, 341, 396, 435, 160, 193, 182, 154],
  // beverly_hills
  [209, 226, 220, 237, 259, 215, 237, 270, 237, 253, 286, 297, 303, 325, 319, 308, 314, 325, 358,  61, 358, 385, 391, 440, 198, 440, 473, 495, 550, 660, 770, 198, 275, 253, 325, 385, 259, 286, 495,  25, 187, 198, 138, 457, 308, 286, 319],
  // pyramids_gardens
  [143, 116, 154, 116, 138, 143, 165, 198,  83, 198, 198, 165, 215, 226, 264, 292, 275, 198, 292, 182, 281, 253, 336, 314, 182, 391, 413, 440, 495, 550, 660, 171, 154, 182, 275, 330, 116, 237, 341, 187,  25, 121, 121, 402, 198, 187, 286],
  // october_gardens
  [242, 231, 253, 220, 198, 248, 253, 303, 171, 297, 330, 226, 292, 286, 303, 336, 363, 253, 352, 182, 407, 308, 391, 369, 121, 462, 468, 495, 550, 660, 715, 264, 209, 303, 319, 358, 242, 330, 396, 198, 121,  25, 182, 583, 363, 237, 336],
  // new_giza
  [149, 171, 165, 154, 182, 154, 182, 209, 138, 209, 237, 209, 226, 253, 264, 248, 286, 237, 308, 110, 292, 286, 369, 358, 165, 413, 413, 440, 495, 550, 660, 171, 198, 193, 264, 303, 182, 248, 435, 138, 121, 182,  25, 418, 220, 209, 275],
  // mv_icity
  [303, 297, 286, 308, 303, 319, 281, 242, 325, 286, 231, 319, 220, 281, 215, 253, 264, 270, 231, 446, 215, 363, 171, 176, 616, 237, 110, 127, 187, 275, 275, 440, 292, 275, 209, 248, 303, 248, 160, 457, 402, 583, 418,  25, 275, 253, 231],
  // zahraa_el_maadi
  [154, 154, 143, 154, 121, 171, 138, 160, 143, 176, 182,  66, 154, 138, 143, 193, 215, 105, 220, 303, 248, 149, 198, 171, 385, 286, 292, 308, 363, 396, 440, 330, 116, 176, 187, 220, 127, 193, 193, 308, 198, 363, 220, 275,  25,  55, 143],
  // carfour_el_maadi
  [143, 138, 138, 149, 116, 165, 127, 143, 143, 171, 165,  88, 143, 116, 116, 182, 253,  44, 198, 275, 253, 138, 182, 149, 319, 292, 292, 303, 369, 468, 385, 270,  88, 138, 171, 220, 110, 176, 182, 286, 187, 237, 209, 253,  55,  25, 143],
  // zahraa_nasr_city
  [165, 160, 154, 171, 165, 198, 143, 121, 198, 165, 121, 182, 121, 132,  66, 143, 193, 116, 143, 380, 154, 275,  99, 149, 418, 193, 193, 215, 275, 358, 385, 319, 154, 165,  94, 127, 165, 132, 154, 319, 286, 336, 275, 231, 143, 143,  25],
];

// ─── Build the nested costs map from the raw matrix ──────────────────────────

function buildCostsMap() {
  const costs = {};
  const ids = REGIONS.map(r => r.id);
  ids.forEach((fromId, i) => {
    costs[fromId] = {};
    ids.forEach((toId, j) => {
      costs[fromId][toId] = RAW_MATRIX[i][j];
    });
  });
  return costs;
}

// ─── Symmetry verification (run once before seeding, check console) ───────────

function verifySym() {
  const ids = REGIONS.map(r => r.id);
  let ok = true;
  for (let i = 0; i < ids.length; i++) {
    for (let j = i + 1; j < ids.length; j++) {
      const ab = RAW_MATRIX[i][j];
      const ba = RAW_MATRIX[j][i];
      if (ab !== ba) {
        console.error(`Asymmetry: ${ids[i]} → ${ids[j]} = ${ab}, reverse = ${ba}`);
        ok = false;
      }
    }
  }
  if (ok) console.log('✅ Matrix verified symmetric (all 1,081 pairs match)');
  return ok;
}

// ─── Main seed function — call from admin dashboard ──────────────────────────
// Requires Firebase 10.x modular SDK already initialised (app, db).

async function seedShippingMatrix(adminPasscode) {
  // Import from same CDN the admin page uses:
  const { doc, setDoc, serverTimestamp } =
    await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js');

  if (!verifySym()) {
    console.error('Seed aborted — symmetry check failed. Fix raw data first.');
    return;
  }

  const payload = {
    regions:    REGIONS,
    costs:      buildCostsMap(),
    version:    1,
    updated_at: serverTimestamp(),
    updated_by: adminPasscode,
  };

  // db must be in scope — the admin page exposes it globally as `window.db`
  await setDoc(doc(window.db, 'shipping_fees', 'matrix'), payload);
  console.log('✅ /shipping_fees/matrix written — 47 regions × 47 = 2,209 cells');
}

// ─── Client-side lookup helper (used by audit UI, zero Firestore reads) ───────

let _cachedMatrix = null;

async function loadMatrix() {
  if (_cachedMatrix) return _cachedMatrix;
  const { doc, getDoc } =
    await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js');
  const snap = await getDoc(doc(window.db, 'shipping_fees', 'matrix'));
  if (!snap.exists()) throw new Error('Matrix not seeded yet');
  _cachedMatrix = snap.data();
  return _cachedMatrix;
}

/**
 * Equivalent of: =INDEX(costs, MATCH(from,rowHeaders,0), MATCH(to,colHeaders,0))
 * Pure in-memory after first load. O(1).
 */
function lookupFee(fromRegionId, toRegionId) {
  if (!_cachedMatrix) throw new Error('Matrix not loaded — await loadMatrix() first');
  const fee = _cachedMatrix.costs?.[fromRegionId]?.[toRegionId];
  if (fee === undefined) throw new Error(`No entry: ${fromRegionId} → ${toRegionId}`);
  return fee;  // EGP integer
}

// ─── Admin: update a single cell (writes both directions for symmetry) ────────

async function updateMatrixCell(fromId, toId, newCost, adminPasscode) {
  const { doc, updateDoc, increment, serverTimestamp } =
    await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js');

  // Read current costs, patch both directions, write back
  const matrix = await loadMatrix();
  const updatedCosts = JSON.parse(JSON.stringify(matrix.costs));
  updatedCosts[fromId][toId] = newCost;
  updatedCosts[toId][fromId] = newCost;   // symmetry

  await updateDoc(doc(window.db, 'shipping_fees', 'matrix'), {
    costs:      updatedCosts,
    version:    increment(1),
    updated_at: serverTimestamp(),
    updated_by: adminPasscode,
  });

  _cachedMatrix = null;   // invalidate cache
  console.log(`Updated: ${fromId} ↔ ${toId} = ${newCost} EGP`);
}

// Export for use in admin page (if using ES modules)
// In a plain script tag, these are global functions.
if (typeof module !== 'undefined') {
  module.exports = { REGIONS, RAW_MATRIX, buildCostsMap, verifySym,
                     seedShippingMatrix, loadMatrix, lookupFee, updateMatrixCell };
}
