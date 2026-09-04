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

  // Public registration tiers, in order. To close a tier once its slots are
  // full, set soldOut: true — it stays visible but shows "Agotado" and can no
  // longer be selected on the registration page. Do not remove a tier that
  // already has registrations tied to it.
  tiers: [
    { id: 'super-early-bird', name: 'Cupo 1', priceUsd: 80,  deadline: 'hasta el 20 de septiembre de 2026', slots: 20, soldOut: false },
    { id: 'early-bird',       name: 'Cupo 2',        priceUsd: 90,  deadline: 'hasta el 10 de octubre de 2026',    slots: 20, soldOut: false, featured: true },
    { id: 'regular',          name: 'Cupo 3',           priceUsd: 100, deadline: 'hasta el 30 de octubre de 2026',    slots: 20, soldOut: false },
    { id: 'last-minute',      name: 'Cupo 4',       priceUsd: 110, deadline: 'hasta el 20 de noviembre de 2026',  slots: 20, soldOut: false },
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

  // Second hidden tier, unlocked the same way with its own code — for special
  // guests who attend the seminar free of charge. specialPractice below is
  // still charged separately even when this code is used. Its distinct `name`
  // is what makes it show up as its own group in the "Precios cupo" summary.
  guestCode: {
    code: 'INVITADO01',
    name: 'Invitado especial',
    priceUsd: 0,
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
    slots: 40,
    soldOut: false,
  },

};

// Fetches live price/slot/soldOut data from the "Precios" and "Inscripciones"
// sheets (via the Apps Script's doGet) and overrides the static tiers above.
// If the request fails for any reason, the static values above are used as-is —
// call this before rendering, but don't let a failure block the page.
SEMINARIO_CONFIG.loadLiveTierData = async function () {
  try {
    const res = await fetch(SEMINARIO_CONFIG.registrationEndpoint, { method: 'GET' });
    const data = await res.json();
    if (!data.ok || !Array.isArray(data.tiers)) throw new Error('respuesta inválida');
    data.tiers.forEach(function (live) {
      // 'clase-especial' en la planilla "Precios" corresponde a specialPractice,
      // no a un tier — se aplica aparte.
      if (live.id === 'clase-especial') {
        const addon = SEMINARIO_CONFIG.specialPractice;
        addon.priceUsd = live.priceUsd;
        addon.slots = live.slots;
        addon.soldOut = addon.soldOut || live.soldOut;
        return;
      }
      // 'access-code' corresponde al Cupo Privado (accessCode), no a un tier público.
      if (live.id === 'access-code') {
        const ac = SEMINARIO_CONFIG.accessCode;
        ac.priceUsd = live.priceUsd;
        ac.slots = live.slots;
        ac.soldOut = ac.soldOut || live.soldOut;
        return;
      }
      // 'guest-code' corresponde al Invitado especial (guestCode), no a un tier público.
      if (live.id === 'guest-code') {
        const gc = SEMINARIO_CONFIG.guestCode;
        gc.priceUsd = live.priceUsd;
        gc.slots = live.slots;
        gc.soldOut = gc.soldOut || live.soldOut;
        return;
      }
      const tier = SEMINARIO_CONFIG.tiers.find(function (t) { return t.id === live.id; });
      if (!tier) return;
      tier.priceUsd = live.priceUsd;
      tier.slots = live.slots;
      tier.soldOut = tier.soldOut || live.soldOut;
    });
  } catch (err) {
    // Sin conexión o el endpoint falló — se mantienen los valores estáticos de arriba.
  }
};
