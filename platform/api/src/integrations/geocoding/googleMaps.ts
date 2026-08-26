import { config } from '../../config.js';
import type { GeocodingConnector, GeoResult, AddressResult, LatLngBounds } from './interface.js';

const GEOCODE_URL = 'https://maps.googleapis.com/maps/api/geocode/json';

export const googleMapsGeocoding: GeocodingConnector = {
  async geocode(address, options) {
    const params = new URLSearchParams({
      address,
      key: config.googleMaps.apiKey,
    });
    if (options?.bounds) {
      const b = options.bounds;
      params.set('bounds', `${b.south},${b.west}|${b.north},${b.east}`);
    }

    const resp = await fetch(`${GEOCODE_URL}?${params}`);
    const json = await resp.json();

    if (json.status !== 'OK' || !json.results?.length) return null;

    const result = json.results[0];
    return {
      lat: result.geometry.location.lat,
      lng: result.geometry.location.lng,
      formattedAddress: result.formatted_address,
    };
  },

  async reverseGeocode(lat, lng) {
    const params = new URLSearchParams({
      latlng: `${lat},${lng}`,
      key: config.googleMaps.apiKey,
    });

    const resp = await fetch(`${GEOCODE_URL}?${params}`);
    const json = await resp.json();

    if (json.status !== 'OK' || !json.results?.length) return null;

    const components = json.results[0].address_components || [];
    const get = (type: string) =>
      components.find((c: any) => c.types.includes(type))?.long_name || '';

    return {
      address: `${get('street_number')} ${get('route')}`.trim(),
      city: get('locality') || get('sublocality'),
      state: get('administrative_area_level_1'),
      zip: get('postal_code'),
      country: get('country'),
    };
  },
};
