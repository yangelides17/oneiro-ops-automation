export interface GeoResult {
  lat: number;
  lng: number;
  formattedAddress: string;
}

export interface AddressResult {
  address: string;
  city: string;
  state: string;
  zip: string;
  country: string;
}

export interface LatLngBounds {
  south: number;
  west: number;
  north: number;
  east: number;
}

export interface GeocodingConnector {
  geocode(address: string, options?: { bounds?: LatLngBounds }): Promise<GeoResult | null>;
  reverseGeocode(lat: number, lng: number): Promise<AddressResult | null>;
}
