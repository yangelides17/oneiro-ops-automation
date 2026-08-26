import { Router } from 'express';
import { googleMapsGeocoding } from '../integrations/geocoding/googleMaps.js';
import { requireRole } from '../middleware/roles.js';

const router = Router();

/** POST /api/geocode — Forward geocode an address. */
router.post('/geocode', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const { address, bounds } = req.body;
  if (!address) return res.status(400).json({ error: 'address required' });

  const result = await googleMapsGeocoding.geocode(address, { bounds });
  if (!result) return res.status(404).json({ error: 'Address not found' });
  res.json(result);
});

/** POST /api/reverse-geocode — Reverse geocode lat/lng (for photo watermark). */
router.post('/reverse-geocode', requireRole('owner', 'admin', 'foreman', 'crew'), async (req, res) => {
  const { lat, lng } = req.body;
  if (lat === undefined || lng === undefined) {
    return res.status(400).json({ error: 'lat and lng required' });
  }

  const result = await googleMapsGeocoding.reverseGeocode(lat, lng);
  if (!result) return res.status(404).json({ error: 'Location not found' });
  res.json(result);
});

export default router;
