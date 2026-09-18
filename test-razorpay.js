require('dotenv').config();
const Razorpay = require('razorpay');

console.log('KEY_ID:', process.env.RAZORPAY_KEY_ID);
console.log('KEY_SECRET loaded:', !!process.env.RAZORPAY_KEY_SECRET);

const rzp = new Razorpay({
  key_id: process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET,
});

rzp.orders.create({ amount: 100, currency: 'INR', receipt: 'test-' + Date.now() })
  .then(o => console.log('SUCCESS:', o))
  .catch(e => console.error('FAILED:', e.error || e));