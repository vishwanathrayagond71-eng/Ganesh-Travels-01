// ============================================================
// utils/osmService.js - OpenStreetMap (Nominatim) Integration
// Handles geocoding, reverse-geocoding, and POI discovery
// ============================================================

const https = require('https');
const { v4: uuidv4 } = require('uuid');

// Helper to make https GET requests with standard User-Agent headers
function httpGet(url) {
  return new Promise((resolve, reject) => {
    const options = {
      headers: {
        'User-Agent': 'GaneshTravelsApp/1.0 (contact: admin@ganeshtravels.com)'
      },
      timeout: 6000 // 6 seconds timeout
    };

    const req = https.get(url, options, (res) => {
      let data = '';
      res.on('data', (chunk) => {
        data += chunk;
      });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(data));
          } catch (e) {
            reject(new Error('Failed to parse JSON response: ' + e.message));
          }
        } else {
          reject(new Error(`OSM request failed with status code ${res.statusCode}`));
        }
      });
    });

    req.on('error', (err) => {
      reject(err);
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request timed out'));
    });
  });
}

// Extract a city or landmark name from Nominatim address
function extractCityName(address) {
  if (!address) return '';
  return address.city || address.town || address.village || address.hamlet || address.municipality || address.city_district || address.state_district || address.suburb || address.county || address.district || '';
}

// 1. Geocode a search string into coordinates and details
async function geocodeAddress(query) {
  try {
    let searchQuery = query.trim();
    
    // Remove any trailing country indicator (like , France or , USA) to prevent user from bypassing India restriction
    searchQuery = searchQuery.replace(/,\s*[a-zA-Z\s]+$/, '');
    searchQuery = `${searchQuery}, India`;

    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(searchQuery)}&format=json&addressdetails=1&limit=1`;
    const results = await httpGet(url);
    
    if (results && results.length > 0) {
      const place = results[0];
      const address = place.address || {};
      const country = address.country || '';
      
      // Strict check: result must be in India
      if (country.toLowerCase().includes('india') || (address.country_code && address.country_code.toLowerCase() === 'in')) {
        const name = place.display_name;
        const lat = parseFloat(place.lat);
        const lng = parseFloat(place.lon);
        const city = extractCityName(address) || place.name || name.split(',')[0] || '';
        const state = address.state || '';
        return { name, lat, lng, city, state, country: 'India' };
      }
    }
    
    return null;
  } catch (err) {
    console.error(`[OSM Geocoding Error] Failed for query "${query}":`, err.message);
    return null;
  }
}

// 2. Reverse geocode coordinates to get address details
async function reverseGeocode(lat, lng) {
  try {
    const url = `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=10`;
    const result = await httpGet(url);
    if (result && result.address) {
      const city = extractCityName(result.address) || result.name || result.display_name.split(',')[0] || '';
      const state = result.address.state || '';
      const country = result.address.country || 'India';
      return { city, state, country, displayName: result.display_name };
    }
    return null;
  } catch (err) {
    console.error(`[OSM Reverse Geocoding Error] Failed for coordinates ${lat}, ${lng}:`, err.message);
    return null;
  }
}

// Curated travel photos for mock/default POIs
const CATEGORY_IMAGES = {
  tourist: [
    'https://images.unsplash.com/photo-1599661046289-e31897846e41?w=600',
    'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
    'https://images.unsplash.com/photo-1476514525535-07fb3b4ae5f1?w=600'
  ],
  temple: [
    'https://images.unsplash.com/photo-1564507592333-c60657eea523?w=600',
    'https://images.unsplash.com/photo-1561361058-c24e01e34f44?w=600',
    'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600'
  ],
  restaurant: [
    'https://images.unsplash.com/photo-1512343879784-a960bf40e7f2?w=600',
    'https://images.unsplash.com/photo-1555396273-367ea4eb4db5?w=600'
  ],
  lodge: [
    'https://images.unsplash.com/photo-1544551763-46a013bb70d5?w=600',
    'https://images.unsplash.com/photo-1566073771259-6a8506099945?w=600'
  ],
  falls: [
    'https://images.unsplash.com/photo-1469474968028-56623f02e42e?w=600',
    'https://images.unsplash.com/photo-1470071459604-3b5ec3a7fe05?w=600'
  ]
};

function getRandomImage(category) {
  const list = CATEGORY_IMAGES[category] || CATEGORY_IMAGES.tourist;
  return list[Math.floor(Math.random() * list.length)];
}

// Generate premium mock POIs in case OSM API is rate-limited or fails
function generateMockPois(cityName, centerLat, centerLng) {
  console.log(`[OSM Fallback] Generating high-quality local POIs for ${cityName}...`);
  const categories = ['tourist', 'temple', 'restaurant', 'lodge', 'falls'];
  const pois = [];

  const localNames = {
    tourist: ['Heritage Palace', 'Eco Tourism Park', 'Historic Fort', 'Crafts Village'],
    temple: ['Shiva Temple', 'Ganesha Temple', 'Sri Devi Temple', 'Hanuman Temple'],
    restaurant: ['Heritage Veg Hotel', 'Flavors Multi-Cuisine', 'Garden Restaurant', 'Grand Kaveri Veg'],
    lodge: ['Comfort Residency & Lodge', 'Pine Hill Resort', 'Transit Inn Suites', 'Kalyan Palace Lodge'],
    falls: ['Milky Waterfalls', 'Green Valley Cascades', 'Scenic Nature Falls']
  };

  categories.forEach(cat => {
    const count = cat === 'falls' ? 1 : 2; // Waterfalls are rarer
    for (let i = 0; i < count; i++) {
      // Add a small random offset within 20km (approx 0.15 degrees)
      const offsetLat = (Math.random() - 0.5) * 0.15;
      const offsetLng = (Math.random() - 0.5) * 0.15;
      const lat = (parseFloat(centerLat) + offsetLat).toFixed(4);
      const lng = (parseFloat(centerLng) + offsetLng).toFixed(4);

      const name = `${cityName} ${localNames[cat][i % localNames[cat].length]}`;
      const rating = (4.0 + Math.random() * 0.9).toFixed(1);

      pois.push({
        id: `poi-mock-${uuidv4().substring(0, 8)}`,
        name,
        category: cat,
        lat,
        lng,
        description: `A highly rated local ${cat} attraction situated in the vibrant vicinity of ${cityName}. Clean surroundings and tourist-friendly environment.`,
        address: `${name}, Near City Center, ${cityName}`,
        rating,
        image: getRandomImage(cat),
        createdAt: new Date().toLocaleString()
      });
    }
  });

  return pois;
}

// Generate premium mock POIs for a specific category
function generateMockPoisForCategory(cityName, centerLat, centerLng, category, count = 3) {
  const localNames = {
    tourist: ['Heritage Palace', 'Eco Tourism Park', 'Historic Fort', 'Crafts Village', 'Nature Garden', 'Historic Museum'],
    temple: ['Shiva Temple', 'Ganesha Temple', 'Sri Devi Temple', 'Hanuman Temple', 'Basaveshwara Temple', 'Durga Temple'],
    restaurant: ['Heritage Veg Hotel', 'Flavors Multi-Cuisine', 'Garden Restaurant', 'Grand Kaveri Veg', 'Highway Inn Dining', 'Local Spice Restaurant'],
    lodge: ['Comfort Residency & Lodge', 'Pine Hill Resort', 'Transit Inn Suites', 'Kalyan Palace Lodge', 'Greenwood Hotel', 'Sunrise Lodge'],
    falls: ['Milky Waterfalls', 'Green Valley Cascades', 'Scenic Nature Falls', 'Hidden Gorge Falls']
  };

  const pois = [];
  const list = localNames[category] || localNames.tourist;
  
  for (let i = 0; i < count; i++) {
    // Add small offset (within 12km)
    const offsetLat = (Math.random() - 0.5) * 0.12;
    const offsetLng = (Math.random() - 0.5) * 0.12;
    const lat = (parseFloat(centerLat) + offsetLat).toFixed(4);
    const lng = (parseFloat(centerLng) + offsetLng).toFixed(4);

    const name = `${cityName} ${list[i % list.length]}`;
    const rating = (4.0 + Math.random() * 0.9).toFixed(1);

    pois.push({
      id: `poi-mock-${uuidv4().substring(0, 8)}`,
      name,
      category,
      lat,
      lng,
      description: `A highly rated local ${category} attraction situated in the vibrant vicinity of ${cityName}. Clean surroundings and tourist-friendly environment.`,
      address: `${name}, Near City Center, ${cityName}`,
      rating,
      image: getRandomImage(category),
      createdAt: new Date().toLocaleString()
    });
  }
  return pois;
}

// 3. Fetch POIs for a city name (either from OSM or via Fallback generator)
async function getPoisForCity(cityName, centerLat, centerLng) {
  try {
    const categories = [
      { osmQuery: 'tourist+attraction', appCat: 'tourist', limit: 4 },
      { osmQuery: 'place_of_worship', appCat: 'temple', limit: 4 },
      { osmQuery: 'restaurant', appCat: 'restaurant', limit: 4 },
      { osmQuery: 'hotel', appCat: 'lodge', limit: 4 },
      { osmQuery: 'waterfall', appCat: 'falls', limit: 2 }
    ];

    const allPois = [];
    const lat = parseFloat(centerLat);
    const lng = parseFloat(centerLng);
    const hasCoords = !isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0;

    // Loop through each category and fetch from Nominatim
    for (const cat of categories) {
      let catPois = [];
      try {
        let url;
        if (hasCoords) {
          // Bounding box of approx. 18km (0.16 degrees) around the center coordinates
          const delta = 0.16;
          const lon1 = lng - delta;
          const lat1 = lat + delta;
          const lon2 = lng + delta;
          const lat2 = lat - delta;
          url = `https://nominatim.openstreetmap.org/search?q=${cat.osmQuery}&viewbox=${lon1},${lat1},${lon2},${lat2}&bounded=1&format=json&limit=${cat.limit}`;
        } else {
          // Fallback to name search if coordinates are missing
          const queryStr = `${cat.osmQuery}+in+${encodeURIComponent(cityName)}`;
          url = `https://nominatim.openstreetmap.org/search?q=${queryStr}&format=json&limit=${cat.limit}`;
        }
        
        // Wait a tiny bit between requests to avoid OSM block
        await new Promise(r => setTimeout(r, 600));

        const results = await httpGet(url);
        
        if (results && results.length > 0) {
          results.forEach((place, index) => {
            const pLat = parseFloat(place.lat);
            const pLng = parseFloat(place.lon);
            if (isNaN(pLat) || isNaN(pLng)) return;

            // Generate clean name
            let name = place.name || place.display_name.split(',')[0];
            if (!name || name.length < 3) name = `${cityName} Local ${cat.appCat}`;

            const rating = (4.2 + (index * 0.1) + Math.random() * 0.3).toFixed(1);

            catPois.push({
              id: `poi-osm-${place.place_id || uuidv4().substring(0, 8)}`,
              name,
              category: cat.appCat,
              lat: String(pLat),
              lng: String(pLng),
              description: `Beautifully preserved local ${cat.appCat} point of interest in ${cityName}. Admired by visitors for its rich local character.`,
              address: place.display_name,
              rating: String(rating > 5 ? 5.0 : rating),
              image: getRandomImage(cat.appCat),
              createdAt: new Date().toLocaleString()
            });
          });
        }
      } catch (err) {
        console.warn(`[OSM Category Warn] Failed for query "${cat.osmQuery}" near "${cityName}":`, err.message);
      }

      // If we couldn't get any POIs for this category from Nominatim, generate mock POIs for this category!
      if (catPois.length === 0) {
        const count = cat.appCat === 'falls' ? 1 : 3;
        catPois = generateMockPoisForCategory(cityName, centerLat, centerLng, cat.appCat, count);
      }
      
      allPois.push(...catPois);
    }

    // Double check: if still empty (unexpected error), trigger complete mock generator
    if (allPois.length === 0) {
      return generateMockPois(cityName, centerLat, centerLng);
    }

    return allPois;
  } catch (err) {
    console.error(`[OSM Sights Error] Failed for city "${cityName}":`, err.message);
    return generateMockPois(cityName, centerLat, centerLng);
  }
}

module.exports = {
  geocodeAddress,
  reverseGeocode,
  getPoisForCity,
  generateMockPois
};
