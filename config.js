/*
 * Shared settings for index.html and inscripcion_seminario_aikido.html.
 * Both pages load this file and render their pricing/contact info from it,
 * so editing a value here updates it everywhere.
 */
const SEMINARIO_CONFIG = {

  contactEmail: 'info@circuloaikikai.com',

  // Google Apps Script Web App URL that receives registration submissions —
  // logs each one to a Sheet and emails a notification. See doPost() in the
  // Apps Script project bound to the "Inscripciones" spreadsheet.
  registrationEndpoint: 'https://script.google.com/macros/s/AKfycby5wimo2vAWHCmvAZ62HsoDEOpjEnAF2nfHdjr-S6K0Knqbf2yrkBxwaPDJhOVdAsvd/exec',

  // ARS shown as an estimate next to each USD price ("≈ $X ARS"). Update this
  // to the current exchange rate — everything recalculates from this one value.
  exchangeRateArsPerUsd: 1550,

  // Public registration tiers, in order. To close a tier once its slots are
  // full, set soldOut: true — it stays visible but shows "Agotado" and can no
  // longer be selected on the registration page. Do not remove a tier that
  // already has registrations tied to it.
  tiers: [
    { id: 'super-early-bird', name: 'Super Early Bird', priceUsd: 80,  deadline: 'hasta el 4 de septiembre 2026',           slots: 20, soldOut: false },
    { id: 'early-bird',       name: 'Early Bird',        priceUsd: 90,  deadline: 'hasta 23 septiembre 2026',       slots: 20, soldOut: false, featured: true },
    { id: 'regular',          name: 'Regular',           priceUsd: 100,  deadline: 'hasta 23 octubre 2026',          slots: 20, soldOut: false },
    { id: 'last-minute',      name: 'Last Minute',       priceUsd: 110, deadline: 'hasta 21 noviembre · en puerta', slots: 20, soldOut: false },
  ],

  // Hidden tier unlocked on the registration page by entering `code`.
  // Same soldOut behavior as the tiers above.
  accessCode: {
    code: 'AIKIKAI01',
    name: 'Cupo Privado',
    priceUsd: 60,
    deadline: 'Solo con código de acceso',
    slots: 10,
    soldOut: false,
  },

  // Optional add-on practice session, offered alongside every tier.
  // Same soldOut behavior as the tiers above — once its own slots fill up,
  // set soldOut: true to stop new selections regardless of which cupo is chosen.
  specialPractice: {
    label: '+ Práctica especial · Viernes 20 de noviembre',
    priceUsd: 30,
    slots: 20,
    soldOut: false,
  },

};
