/**
 * QuickBooks Online Service-item IDs per pricing group.
 *
 * Each of the 5 IDs is the QBO internal Item Id captured from the QBO
 * UI after the Service items are manually created in the target QBO
 * company (see docs/quickbooks_integration.md for the setup steps).
 *
 * Why not env vars: these are not secrets, they rarely change, and
 * keeping them in code makes diffs explicit. To swap sandbox → prod,
 * edit this file and redeploy.
 *
 * Boot-time assertion (assertQbItemsConfigured) is called from
 * server.js so a misconfig fails the webapp on startup instead of
 * silently failing the first invoice generation.
 */

export const QB_ITEMS = Object.freeze({
  line4:         '148',  // Extruded Thermo Lines
  line12:        '149',  // Extruded Thermo Crosswalks/Stop Lines
  preformed:     '150',  // Preformed L&S
  extruded:      '151',  // Extruded L&S
  color_surface: '152',  // Color Surface
})

export function assertQbItemsConfigured() {
  // Only fail boot if QB integration is meant to be live (i.e. all
  // required env vars are present). Lets a dev run the webapp without
  // QB credentials configured.
  const enabled = process.env.QB_CLIENT_ID
               && process.env.QB_CLIENT_SECRET
               && process.env.QB_REALM_ID
  if (!enabled) return

  const missing = Object.entries(QB_ITEMS)
    .filter(([, id]) => !id || !String(id).trim())
    .map(([group]) => group)
  if (missing.length > 0) {
    throw new Error(
      `QB_ITEMS missing for: ${missing.join(', ')}. ` +
      `Edit webapp/server/qbItems.js with the Service-item IDs from QBO.`
    )
  }
}
