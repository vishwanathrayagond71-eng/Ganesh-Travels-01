// ============================================================
// chat.js - Simple Chatbot Script
// Pre-programmed responses for common travel questions
// ============================================================

// Chatbot knowledge base
const chatResponses = {
  'best destinations in india': '🇮🇳 Our top Indian destinations are:\n• Taj Mahal, Agra\n• Kerala Backwaters\n• Ladakh\n• Goa Beaches\n• Manali, Himachal\n• Jaipur, Rajasthan\nCheck our <a href="/destinations" style="color:#e9c46a">Destinations page</a> for more!',
  
  'what are your tour packages': '📦 We offer:\n• Family Packages\n• Honeymoon Specials\n• Adventure Trips\n• Solo Getaways\n• Group Tours\n• International Tours\nVisit our <a href="/packages" style="color:#e9c46a">Packages page</a> to explore!',
  
  'how to book a tour': '📅 Booking is easy! Just:\n1. Go to our <a href="/booking" style="color:#e9c46a">Booking page</a>\n2. Select your destination & dates\n3. Enter your details\n4. Click "Confirm Booking"\nWe\'ll confirm via email within 24 hours! ✅',
  
  'cancellation policy': `❌ Our cancellation policy:\n• Free cancellation within 24 hours\n• 7+ days before: 75% refund\n• Within 7 days: No refund\nContact us for special cases at ${window.contactEmail || 'xyz7@gmail.com'}`,
  
  'price': '💰 Our prices start from:\n• Budget trips: ₹7,999/person\n• Family packages: ₹24,999\n• Honeymoon: ₹34,999\n• International: ₹54,999+\nCheck our <a href="/packages" style="color:#e9c46a">Packages</a> for exact pricing!',
  
  'contact': `📞 Reach us at:\n• Phone: ${window.contactPhone || '+91 98765 43210'}\n• Email: ${window.contactEmail || 'xyz7@gmail.com'}\n• Office: Ganesh Travels, Vidyagiri, Bagalkote, Karnataka\n• Hours: Mon-Sat 9AM-7PM\nOr visit our <a href="/contact" style="color:#e9c46a">Contact page</a>!`,
  
  'goa': '🌊 Goa is fantastic! Our Goa package includes:\n• Boutique hotel stay\n• Beach activities\n• Scooter rental\n• City tour\n• Starting from ₹18,999\n<a href="/booking?destination=Goa Beaches" style="color:#e9c46a">Book Goa Now →</a>',
  
  'ladakh': '🏔️ Ladakh is breathtaking! Package includes:\n• Camp & hotel stays\n• Bike rental option\n• Nubra Valley & Pangong Lake\n• Expert adventure guide\n• Starting from ₹44,999\n<a href="/booking?destination=Ladakh" style="color:#e9c46a">Book Ladakh Now →</a>',
  
  'kerala': '🌿 Kerala — God\'s Own Country! Includes:\n• Houseboat stay in Alleppey\n• Munnar hill station visit\n• Thekkady wildlife\n• Starting from ₹34,999\n<a href="/booking?destination=Kerala Backwaters" style="color:#e9c46a">Book Kerala Now →</a>',
  
  'international': '🌍 Popular international destinations:\n• Bali, Indonesia — ₹54,999\n• Dubai, UAE — ₹79,999\n• Paris, France — ₹89,999\n• Maldives — ₹1,29,999\n<a href="/destinations" style="color:#e9c46a">Explore International Tours →</a>',

  'hello': `👋 Hello! Welcome to Ganesh Travels! I'm your travel assistant. How can I help you plan your dream vacation? You can ask me about destinations, packages, pricing, or how to book! 😊`,
  
  'hi': '👋 Hi there! Great to meet you! I\'m your virtual travel guide. Ask me anything about our tours, destinations, or packages! 🌍',
  
  'thanks': '😊 You\'re welcome! Feel free to ask if you have more questions. Happy travels! ✈️',
  
  'help': '🤝 I can help you with:\n• Finding destinations\n• Tour package information\n• Pricing queries\n• Booking process\n• Cancellation policy\n• Contact information\nWhat would you like to know?'
};

// Default response for unrecognized questions
const defaultResponse = `🤔 I'm not sure about that specific question. For personalized assistance, please:\n• <a href='/contact' style='color:#e9c46a'>Contact our team</a>\n• Call us: ${window.contactPhone || '+91 98765 43210'}\n• Email: ${window.contactEmail || 'xyz7@gmail.com'}\nOur experts are available Mon-Sat 9AM-7PM! 😊`;

// Toggle chatbot open/closed
function toggleChatbot() {
  const widget = document.getElementById('chatbotWidget');
  const badge = document.querySelector('.chatbot-badge');
  if (widget) {
    widget.classList.toggle('open');
    // Hide badge when opened
    if (widget.classList.contains('open') && badge) {
      badge.style.display = 'none';
    }
  }
}

// Send a message from the user
function sendChat() {
  const input = document.getElementById('chatInput');
  if (!input) return;
  const message = input.value.trim();
  if (!message) return;
  addMessage(message, 'user');
  input.value = '';
  // Simulate bot thinking delay
  setTimeout(() => {
    const response = getBotResponse(message);
    addBotMessage(response);
  }, 600);
}

// Send a pre-set quick question
function sendChatQuestion(question) {
  addMessage(question, 'user');
  setTimeout(() => {
    const response = getBotResponse(question);
    addBotMessage(response);
  }, 400);
}

// Add a message to the chat UI
function addMessage(text, sender) {
  const messages = document.getElementById('chatMessages');
  if (!messages) return;
  const div = document.createElement('div');
  div.className = `chat-msg ${sender === 'user' ? 'user-msg' : ''}`;
  if (sender === 'bot') {
    div.innerHTML = `<i class="fas fa-robot"></i><span>${text}</span>`;
  } else {
    div.innerHTML = `<span>${text}</span><i class="fas fa-user"></i>`;
  }
  messages.appendChild(div);
  messages.scrollTop = messages.scrollHeight;
}

function addBotMessage(text) {
  const messages = document.getElementById('chatMessages');
  if (!messages) return;
  const div = document.createElement('div');
  div.className = 'chat-msg bot-msg';
  div.innerHTML = `<i class="fas fa-robot"></i><span style="white-space:pre-line">${text}</span>`;
  messages.appendChild(div);
  messages.scrollTop = messages.scrollHeight;
}

// Find the best response from the knowledge base
function getBotResponse(message) {
  const lower = message.toLowerCase().trim();
  
  // Check for keyword matches
  for (const [key, response] of Object.entries(chatResponses)) {
    const keywords = key.split(' ');
    if (keywords.some(word => lower.includes(word)) || lower.includes(key)) {
      return response;
    }
  }
  
  // Specific keyword checks
  if (lower.includes('book') || lower.includes('reserv')) return chatResponses['how to book a tour'];
  if (lower.includes('cancel') || lower.includes('refund')) return chatResponses['cancellation policy'];
  if (lower.includes('price') || lower.includes('cost') || lower.includes('rate') || lower.includes('₹')) return chatResponses['price'];
  if (lower.includes('contact') || lower.includes('phone') || lower.includes('email')) return chatResponses['contact'];
  if (lower.includes('package') || lower.includes('tour')) return chatResponses['what are your tour packages'];
  if (lower.includes('destination') || lower.includes('place') || lower.includes('travel')) return chatResponses['best destinations in india'];
  if (lower.includes('goa') || lower.includes('beach')) return chatResponses['goa'];
  if (lower.includes('ladakh') || lower.includes('himalaya') || lower.includes('mountain')) return chatResponses['ladakh'];
  if (lower.includes('kerala') || lower.includes('backwater')) return chatResponses['kerala'];
  if (lower.includes('international') || lower.includes('abroad') || lower.includes('bali') || lower.includes('dubai')) return chatResponses['international'];
  
  return defaultResponse;
}
