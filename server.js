/**
 * ============================================================================
 *  MAPMYCAREER360 - EDUCATION CONSULTANT WEBSITE
 *  BACKEND - SINGLE FILE SERVER (server.js)
 * ----------------------------------------------------------------------------
 *  Per client request: everything (DB models, auth, all API routes) lives in
 *  this one file - NO separate routes/ or middleware/ folders/files.
 *
 *  Covers every feature from "Education_Consultant_Website_Quotation_final.pdf":
 *   - Student registration / login / password reset / profile
 *   - Career Assessment (paid + free-for-subscribers) + results
 *   - Subscription plans (Monthly / Quarterly / Yearly)
 *   - Payment Gateway integration hook (Razorpay/PhonePe/Stripe) + invoices
 *   - Student Dashboard data (profile, subscription, assessments, payments)
 *   - KCET Predictor
 *   - PGCET Predictor
 *   - Referral system (assessment/subscription referrals)
 *   - College Admission Referral (separate, can be hidden from student view)
 *   - College Compare
 *   - Admin Panel: dashboard, students, subscriptions, assessments, payments,
 *     referrals, college referrals, KCET/PGCET cutoff data, colleges, courses,
 *     slider/banners, website pages/content, reports, settings
 * ============================================================================
 */

require('dotenv').config();

console.log('RAZORPAY_KEY_ID:', process.env.RAZORPAY_KEY_ID);
console.log(
  'RAZORPAY_KEY_SECRET loaded:',
  !!process.env.RAZORPAY_KEY_SECRET
);

const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const dns = require('dns'); 
const { execFileSync } = require('child_process');

dns.setServers(['8.8.8.8', '1.1.1.1']);

const app = express();

/* ============================================================================
 *  BASIC CONFIG
 * ==========================================================================*/
const PORT = process.env.PORT || 5000;
const NODE_ENV = process.env.NODE_ENV || 'development';
const JWT_SECRET = process.env.JWT_SECRET || 'dev_only_change_me';
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '7d';
const MONGO_URI = process.env.MONGO_URI 

// Connection pool sizing - tune via .env as traffic grows. Mongoose/MongoDB driver
// reuses this pool for every request, so we do NOT open a new connection per request.
const DB_MAX_POOL_SIZE = Number(process.env.DB_MAX_POOL_SIZE || 20);
const DB_MIN_POOL_SIZE = Number(process.env.DB_MIN_POOL_SIZE || 2);

// Hard ceilings so no endpoint can accidentally return an unbounded dataset
const DEFAULT_PAGE_SIZE = 20;
const MAX_PAGE_SIZE = 100;
const MAX_EXPORT_ROWS = 5000;
const MAX_PREDICTOR_RESULTS = 60;


const PREDICTOR_ONE_TIME_FEE = 99; // set your price
// Dev/demo only - lets the KCET/PGCET predictors (and, separately, the
// Career Assessment page) bypass the payment/subscription check entirely.
// Gated server-side so a client can never enable this themselves by sending
// a flag - it only works if this env var is explicitly set to 'true' on
// THIS server. Never set this in a production .env file.
const ALLOW_DEV_SKIP_PAYMENT = process.env.ALLOW_DEV_SKIP_PAYMENT === 'true';

// Allowed origins - localhost for dev, mapmycareer360.com (+ www) for production.
// You can also add a raw server IP (e.g. 196.x.x.x) via CLIENT_ORIGINS in .env
const CLIENT_ORIGINS = (process.env.CLIENT_ORIGINS ||
  'http://localhost:3000,https://mapmycareer360.com,https://www.mapmycareer360.com'
).split(',').map((o) => o.trim());

app.use(cors({
  origin: function (origin, callback) {
    if (!origin) return callback(null, true); // allow server-to-server / curl / mobile apps
    if (NODE_ENV !== 'production') return callback(null, true); // relaxed in dev
    if (CLIENT_ORIGINS.includes(origin)) return callback(null, true);
    return callback(new Error('Not allowed by CORS: ' + origin));
  },
  credentials: true,
}));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

/* ============================================================================
 *  UPLOADS (slider/banner images, college images, profile photos)
 * ==========================================================================*/
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
app.use('/uploads', express.static(uploadsDir));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  },
});
const upload = multer({ storage, limits: { fileSize: 5 * 1024 * 1024 } });

/* ---- KYC documents: stored privately, never under /uploads ---- */
const kycDir = path.join(__dirname, 'private-uploads', 'kyc');
if (!fs.existsSync(kycDir)) fs.mkdirSync(kycDir, { recursive: true });

const KYC_DOC_TYPES = ['passbook', 'panCard', 'aadhaarCard'];
const KYC_DOC_LABELS = { passbook: 'bank passbook', panCard: 'PAN card', aadhaarCard: 'Aadhaar card' };
const KYC_MIME_EXT = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'application/pdf': '.pdf' };

const kycUpload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, kycDir),
    filename: (req, file, cb) => {
      // extension comes from the checked mime type, not from the user's file name
      cb(null, `${req.user.id}-${file.fieldname}-${Date.now()}${KYC_MIME_EXT[file.mimetype]}`);
    },
  }),
  limits: { fileSize: 5 * 1024 * 1024, files: 3 },
  fileFilter: (req, file, cb) => {
    if (KYC_MIME_EXT[file.mimetype]) return cb(null, true);
    cb(new Error('Only JPG, PNG, WEBP or PDF files are allowed'));
  },
}).fields(KYC_DOC_TYPES.map((name) => ({ name, maxCount: 1 })));

function runKycUpload(req, res, next) {
  kycUpload(req, res, (err) => {
    if (!err) return next();
    const message = err.code === 'LIMIT_FILE_SIZE' ? 'Each file must be 5 MB or smaller' : err.message;
    res.status(400).json({ success: false, message });
  });
}


const Razorpay = require('razorpay');
const crypto = require('crypto');

const razorpay = new Razorpay({
  key_id: process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET,
});

/* ============================================================================
 *  DATABASE CONNECTION
 * ==========================================================================*/
mongoose
  .connect(MONGO_URI, {
    maxPoolSize: DB_MAX_POOL_SIZE, // reused across all requests - NOT one connection per request
    minPoolSize: DB_MIN_POOL_SIZE,
    serverSelectionTimeoutMS: 8000,
  })
  .then(() => console.log(`[MongoDB] connected ->  (pool ${DB_MIN_POOL_SIZE}-${DB_MAX_POOL_SIZE})`))
  .catch((err) => console.error('[MongoDB] connection error:', err.message));

mongoose.set('strictQuery', true);

/* ============================================================================
 *  SCHEMAS / MODELS  (all inline - no separate models/ folder)
 * ==========================================================================*/
const { Schema, model, Types } = mongoose;

// ---- Student ----
const studentSchema = new Schema({
  fullName: { type: String, required: true },
  email: { type: String, required: true, unique: true, lowercase: true, trim: true },
  phone: { type: String, required: true },
  password: { type: String, required: true },
  profilePhoto: { type: String, default: '' },
  gender: { type: String, default: '' },
  dob: { type: Date },
  address: { type: String, default: '' },
  referralCode: { type: String, unique: true }, // this student's own shareable code
  referredBy: { type: String, default: '' },    // referral code that brought this student in
  // Bank/UPI details needed to actually pay out referral earnings. Status
  // moves not_submitted -> pending (on submit) -> verified/rejected (admin).
  kyc: {
    accountHolderName: { type: String, default: '' },
    bankName: { type: String, default: '' },
    accountNumber: { type: String, default: '' },
    ifsc: { type: String, default: '' },
    panNumber: { type: String, default: '' },
    upiId: { type: String, default: '' },
    status: { type: String, enum: ['not_submitted', 'pending', 'verified', 'rejected'], default: 'not_submitted' },
    submittedAt: { type: Date },
    verifiedAt: { type: Date },
    rejectionReason: { type: String, default: '' },
        // Private file names (stored in /private-uploads/kyc, never served publicly)
    documents: {
      passbook: { type: String, default: '' },
      panCard: { type: String, default: '' },
      aadhaarCard: { type: String, default: '' },
    },
    reviewedBy: { type: Schema.Types.ObjectId, ref: 'Admin' },
  },
  resetToken: { type: String },
  resetTokenExpiry: { type: Date },
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

// Indexes: email/referralCode already unique-indexed via schema; add createdAt for
// admin listing sort and referredBy for referral lookups.
studentSchema.index({ createdAt: -1 });
studentSchema.index({ referredBy: 1 });

const Student = model('Student', studentSchema);

// ---- Admin ----
const adminSchema = new Schema({
  name: { type: String, default: 'Administrator' },
  email: { type: String, required: true, unique: true, lowercase: true },
  password: { type: String, required: true },
  role: { type: String, default: 'admin' }, // admin | superadmin
}, { timestamps: true });

const Admin = model('Admin', adminSchema);

// ---- Subscription Plan (admin configurable) ----
const subscriptionPlanSchema = new Schema({
  name: { type: String, required: true }, // Monthly / Quarterly / Yearly
  durationInDays: { type: Number, required: true },
  price: { type: Number, required: true },
  features: [{ type: String }],
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

const SubscriptionPlan = model('SubscriptionPlan', subscriptionPlanSchema);

// ---- Student Subscription (purchase record) ----
const subscriptionSchema = new Schema({
  student: { type: Schema.Types.ObjectId, ref: 'Student', required: true },
  plan: { type: Schema.Types.ObjectId, ref: 'SubscriptionPlan', required: true },
  startDate: { type: Date, default: Date.now },
  endDate: { type: Date, required: true },
  status: { type: String, enum: ['active', 'expired', 'cancelled'], default: 'active' },
  amountPaid: { type: Number, required: true },
  paymentId: { type: Schema.Types.ObjectId, ref: 'Payment' },
}, { timestamps: true });

// Compound index supports "find this student's active subscription" lookups used
// on nearly every authenticated page (assessment access, dashboard, predictors gate).
subscriptionSchema.index({ student: 1, status: 1, endDate: -1 });
subscriptionSchema.index({ createdAt: -1 });

const Subscription = model('Subscription', subscriptionSchema);

// ---- Predictor Lead (captures every KCET/PGCET enquiry, paid or not) ----
const predictorLeadSchema = new Schema({
  name: { type: String, required: true },
  phone: { type: String, required: true },
  email: { type: String, default: '' },
  examType: { type: String, enum: ['kcet', 'pgcet'], required: true },
  rank: { type: Number, required: true },
  category: { type: String, required: true },
  course: { type: String, default: '' },
  student: { type: Schema.Types.ObjectId, ref: 'Student' }, // set if logged in
  accessType: { type: String, enum: ['subscription', 'one-time', 'dev-skip'], required: true },
  paymentId: { type: Schema.Types.ObjectId, ref: 'Payment' }, // for one-time
  resultsShown: { type: Number, default: 0 }, // how many colleges were returned
}, { timestamps: true });

predictorLeadSchema.index({ examType: 1, createdAt: -1 });
predictorLeadSchema.index({ phone: 1 });

const PredictorLead = model('PredictorLead', predictorLeadSchema);




// ---- Career Assessment (questions bank + attempts) ----
// Each option now carries small per-career-field weights instead of a single
// correct answer, e.g. { label: 'Solving math puzzles', weights: { 'Engineering & Technology': 3 } }.
// This mirrors exactly what the multi-section Career Assessment frontend expects,
// so admin-authored questions can replace its local sample banks with no extra mapping.
const assessmentQuestionSchema = new Schema({
  question: { type: String, required: true },
  options: [{
    label: { type: String, required: true },
    weights: { type: Schema.Types.Mixed, default: {} }, // { 'Engineering & Technology': 3, ... }
  }],
  category: { type: String, enum: ['interest', 'aptitude', 'personality', 'adaptive'], default: 'interest' },
  // Only set for category:'adaptive' - which leading career field this question is meant to sharpen.
  adaptiveForField: { type: String, default: '' },
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

assessmentQuestionSchema.index({ category: 1, isActive: 1 });

const AssessmentQuestion = model('AssessmentQuestion', assessmentQuestionSchema);

// ---- Career Database (admin-managed "career field" records) ----
// Mirrors CAREER_DB in the frontend's CareerAssessment.jsx. Once populated,
// point the frontend at GET /api/career-fields instead of its local fallback -
// this is the data source behind the "admin assessment builder" / "career
// database" advanced-plan features.
const careerFieldSchema = new Schema({
  name: { type: String, required: true, unique: true }, // e.g. 'Engineering & Technology'
  stream: { type: String, default: '' },
  careers: [{
    name: { type: String, required: true },
    blurb: { type: String, default: '' },
  }],
  courses: [{ type: String }],
  roadmap: [{
    stage: { type: String, required: true },
    detail: { type: String, default: '' },
  }],
  workStyle: { type: String, default: '' },
  growth: { type: String, default: '' },
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

const CareerField = model('CareerField', careerFieldSchema);

// ---- Career Values (used in the "Career Values" assessment step) ----
const careerValueSchema = new Schema({
  key: { type: String, required: true, unique: true }, // e.g. 'creativity'
  label: { type: String, required: true },              // e.g. 'Creativity & Expression'
  nudge: { type: Schema.Types.Mixed, default: {} },      // { 'Design & Creative Media': 2, ... }
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

const CareerValue = model('CareerValue', careerValueSchema);

const assessmentResultSchema = new Schema({
  student: { type: Schema.Types.ObjectId, ref: 'Student', required: true },

  // Legacy fields - kept so assessment records taken before the multi-section
  // rewrite still read back correctly. New submissions leave these empty and
  // populate the fields below instead.
  answers: [{ questionId: Schema.Types.ObjectId, selectedOptionIndex: Number }],
  scoreSummary: { type: Schema.Types.Mixed, default: {} }, // e.g. { Science: 80, Commerce: 40, Arts: 60 }

  // Multi-section assessment fields (current frontend)
  studentInfo: { type: Schema.Types.Mixed, default: {} },     // { fullName, className, schoolName, city }
  academicProfile: { type: Schema.Types.Mixed, default: {} }, // { board, currentStream, favoriteSubject, marksBand }
  careerValues: [{ type: String }],                            // selected value keys, e.g. ['creativity','learning']
  scores: { type: Schema.Types.Mixed, default: {} },           // { 'Engineering & Technology': 8, ... }
  result: { type: Schema.Types.Mixed, default: {} },           // { archetype, topField, secondField, top10Careers, top3Careers, stream, courses, roadmap, strengths, growthAreas }

  recommendedStreams: [{ type: String }],
  isFreeViaSubscription: { type: Boolean, default: false },
  paymentId: { type: Schema.Types.ObjectId, ref: 'Payment' },
  attemptedAt: { type: Date, default: Date.now },
}, { timestamps: true });

assessmentResultSchema.index({ student: 1, createdAt: -1 });

const AssessmentResult = model('AssessmentResult', assessmentResultSchema);

// ---- Promo Codes ----
const promoCodeSchema = new Schema({
  code: { type: String, required: true, unique: true, uppercase: true, trim: true },
  description: { type: String, default: '' },
  discountType: { type: String, enum: ['percent', 'flat'], required: true },
  discountValue: { type: Number, required: true, min: 0 },
  maxDiscount: { type: Number, default: 0 },           // cap for percent codes, 0 = no cap
  applicablePlans: [{ type: Schema.Types.ObjectId, ref: 'SubscriptionPlan' }], // empty = all plans
  usageLimit: { type: Number, default: 0 },            // total redemptions, 0 = unlimited
  usedCount: { type: Number, default: 0 },
  perStudentLimit: { type: Number, default: 1 },       // 0 = unlimited per student
  validFrom: { type: Date },
  validUntil: { type: Date },
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

const PromoCode = model('PromoCode', promoCodeSchema);

const promoRedemptionSchema = new Schema({
  promo: { type: Schema.Types.ObjectId, ref: 'PromoCode', required: true },
  student: { type: Schema.Types.ObjectId, ref: 'Student', required: true },
  payment: { type: Schema.Types.ObjectId, ref: 'Payment' },
}, { timestamps: true });
promoRedemptionSchema.index({ promo: 1, student: 1 });

const PromoRedemption = model('PromoRedemption', promoRedemptionSchema);




// ---- Payments ----
const paymentSchema = new Schema({
  student: { type: Schema.Types.ObjectId, ref: 'Student', required: true },

    plan: { type: Schema.Types.ObjectId, ref: 'SubscriptionPlan' },
  promo: { type: Schema.Types.ObjectId, ref: 'PromoCode' },
  promoCode: { type: String, default: '' },
  originalAmount: { type: Number },
  discountAmount: { type: Number, default: 0 },
  // find this in paymentSchema and update it:
purpose: { type: String, enum: ['assessment', 'subscription', 'predictor', 'other'], required: true },
  amount: { type: Number, required: true },
  gateway: { type: String, enum: ['razorpay', 'phonepe', 'stripe'], default: 'razorpay' },
  gatewayOrderId: { type: String },
  gatewayPaymentId: { type: String },
  status: { type: String, enum: ['created', 'success', 'failed'], default: 'created' },
  invoiceNumber: { type: String, unique: true, sparse: true },
}, { timestamps: true });

paymentSchema.index({ student: 1, createdAt: -1 });
paymentSchema.index({ status: 1 });

const Payment = model('Payment', paymentSchema);

// ---- Referral system (Career Assessment / Subscription referrals) ----
const referralSchema = new Schema({
  referringStudent: { type: Schema.Types.ObjectId, ref: 'Student', required: true },
  referralCode: { type: String, required: true },
  referredName: { type: String },
  referredEmail: { type: String },
  referredStudent: { type: Schema.Types.ObjectId, ref: 'Student' },
  type: { type: String, enum: ['assessment', 'subscription'], required: true },
  status: { type: String, enum: ['pending', 'converted'], default: 'pending' },
}, { timestamps: true });

referralSchema.index({ referringStudent: 1, createdAt: -1 });
referralSchema.index({ referralCode: 1 });
referralSchema.index({ status: 1 });

const Referral = model('Referral', referralSchema);

// ---- College Admission Referral (separate & can hide college details) ----
const collegeReferralSchema = new Schema({
  referralLinkCode: { type: String, required: true, unique: true },
  createdByAdmin: { type: Boolean, default: true },
  ownerName: { type: String }, // e.g. referring counselor / agent
  ownerContact: { type: String },
  college: { type: Schema.Types.ObjectId, ref: 'College' },
  hideCollegeFromReferredView: { type: Boolean, default: true }, // "college reference can remain hidden"
  enquiries: [{
    name: String,
    phone: String,
    email: String,
    message: String,
    createdAt: { type: Date, default: Date.now },
  }],
  status: { type: String, enum: ['open', 'closed'], default: 'open' },
}, { timestamps: true });

collegeReferralSchema.index({ status: 1, createdAt: -1 });

const CollegeReferral = model('CollegeReferral', collegeReferralSchema);

// ---- Colleges & Courses ----
// Single unified College model - powers the one college hub page (browse,
// filter, compare) plus the college detail page and the admin "Manage
// Colleges" screen. Every field the hub/compare UI shows lives here so the
// frontend never has to stitch data from more than one place.
const collegeSchema = new Schema({
  name: { type: String, required: true },
  code: { type: String },
  location: { type: String },          // city, e.g. "Bangalore"
  state: { type: String, default: '' },
  type: { type: String, enum: ['Government', 'Government-Aided', 'Private', 'Deemed', 'Autonomous'], default: 'Private' },
  establishedYear: { type: Number },
  affiliatedUniversity: { type: String, default: '' },

  logo: { type: String, default: '' },     // square logo shown on cards
  image: { type: String, default: '' },    // banner / cover photo
  gallery: [{ type: String }],             // extra campus photos

  about: { type: String, default: '' },
  description: { type: String, default: '' }, // short one-liner used on cards

  rating: { type: Number, default: 0, min: 0, max: 5 },
  ranking: { type: Number },

  // Accreditation, e.g. ["NAAC A+", "NBA", "AICTE Approved", "UGC Approved"]
  accreditations: [{ type: String }],
  // Specializations / streams offered, e.g. ["Marketing", "Finance", "HR", "Operations", "Analytics"]
  specializations: [{ type: String }],
    // Programs ticked in admin, e.g. ["B.E.", "MBA", "PGDM"]
  programs: [{ type: String }],
  // collegeSchema.index({ programs: 1 });

    // Bangalore zone for the "Area" filter
  area: { type: String, enum: ['', 'North', 'South', 'East', 'West', 'Central'], default: '' },
  // Study modes offered - powers the Full Time / Part Time / Online filter
  courseTypes: [{ type: String, enum: ['Full Time', 'Part Time', 'Online'] }],

  coursesOffered: [{ type: Schema.Types.ObjectId, ref: 'Course' }],

  fees: {
    tuitionAnnual: { type: Number, default: 0 },   // per-year tuition
    totalCourse: { type: Number, default: 0 },      // full course fee
    applicationFee: { type: Number, default: 0 },
    annual: { type: Number, default: 0 },           // kept for backward compatibility with older data
  },

  placements: {
    highestPackage: { type: Number, default: 0 },   // in ₹ Lakhs Per Annum (LPA)
    averagePackage: { type: Number, default: 0 },   // in LPA
    placementPercentage: { type: Number, default: 0 },
    topRecruiters: [{ type: String }],
  },

  hostel: {
    available: { type: Boolean, default: false },
    twoShare: { available: { type: Boolean, default: false }, feesPerYear: { type: Number, default: 0 } },
    threeShare: { available: { type: Boolean, default: false }, feesPerYear: { type: Number, default: 0 } },
    fourShare: { available: { type: Boolean, default: false }, feesPerYear: { type: Number, default: 0 } },
  },

  facilities: [{ type: String }],   // amenities: Library, Labs, Sports Complex, Wi-Fi...

  contact: {
    phone: { type: String, default: '' },
    email: { type: String, default: '' },
    website: { type: String, default: '' },
  },
  brochureUrl: { type: String, default: '' },

  featured: { type: Boolean, default: false },
}, { timestamps: true });

collegeSchema.index({ name: 1 });
collegeSchema.index({ type: 1 });
collegeSchema.index({ location: 1 });
collegeSchema.index({ specializations: 1 });
collegeSchema.index({ programs: 1 }); 
collegeSchema.index({ featured: 1, ranking: 1 });

const College = model('College', collegeSchema);

const courseSchema = new Schema({
  name: { type: String, required: true }, // e.g. Computer Science Engineering
  level: { type: String, enum: ['UG', 'PG', 'Diploma'], default: 'UG' },
  durationYears: { type: Number, default: 4 },
}, { timestamps: true });

const Course = model('Course', courseSchema);

// ---- KCET / PGCET Cutoff data (uploaded by admin, used by predictors) ----
const kcetCutoffSchema = new Schema({
  year: { type: Number, required: true },
  college: { type: Schema.Types.ObjectId, ref: 'College', required: true },
  course: { type: Schema.Types.ObjectId, ref: 'Course', required: true },
  category: { type: String, required: true }, // GM, 2A, 2B, 3A, 3B, SC, ST, 1G, ...
  round: { type: String, default: 'Round 1' },
  cutoffRank: { type: Number, required: true },
  is371J: { type: Boolean, default: false }, // 371J / Hyderabad-Karnataka reservation
}, { timestamps: true });

// This compound index mirrors the exact filter+sort used by /api/predictors/kcet
// (category + cutoffRank range, sorted by cutoffRank) so predictor queries stay
// index-covered even once cutoff data grows into the hundreds of thousands of rows.
kcetCutoffSchema.index({ category: 1, cutoffRank: 1 });
kcetCutoffSchema.index({ year: 1, category: 1, cutoffRank: 1 });
kcetCutoffSchema.index({ college: 1 });
kcetCutoffSchema.index({ course: 1 });

const KcetCutoff = model('KcetCutoff', kcetCutoffSchema);

const pgcetCutoffSchema = new Schema({
  year: { type: Number, required: true },
  college: { type: Schema.Types.ObjectId, ref: 'College', required: true },
  course: { type: Schema.Types.ObjectId, ref: 'Course', required: true },
  category: { type: String, required: true },
  collegeType: { type: String, enum: ['Government', 'Private', 'Autonomous'], default: 'Private' },
  cutoffRank: { type: Number, required: true },
}, { timestamps: true });

pgcetCutoffSchema.index({ category: 1, cutoffRank: 1 });
pgcetCutoffSchema.index({ year: 1, category: 1, cutoffRank: 1 });
pgcetCutoffSchema.index({ college: 1 });
pgcetCutoffSchema.index({ course: 1 });

const PgcetCutoff = model('PgcetCutoff', pgcetCutoffSchema);

// ---- Home page slider / banners ----
// ---- Home page slider / banners ----
const sliderSchema = new Schema({
  title: { type: String, default: '' },
  subtitle: { type: String, default: '' },
  image: { type: String, required: true },
  ctaText: { type: String, default: 'Take Free Career Assessment' },
  ctaLink: { type: String, default: '/career-assessment' },
  imageOnly: { type: Boolean, default: false }, // NEW: image-only slide, no text/link overlay
  order: { type: Number, default: 0 },
  isActive: { type: Boolean, default: true },
}, { timestamps: true });

const Slider = model('Slider', sliderSchema);

// ---- Website pages / general content (About, Services, Contact, etc.) ----
const pageContentSchema = new Schema({
  slug: { type: String, required: true, unique: true }, // e.g. 'home', 'about-us', 'services'
  title: { type: String, default: '' },
  content: { type: Schema.Types.Mixed, default: {} }, // flexible JSON block editable from admin
}, { timestamps: true });

const PageContent = model('PageContent', pageContentSchema);

// ---- Contact messages ----
const contactMessageSchema = new Schema({
  name: String,
  email: String,
  phone: String,
  message: String,
  status: { type: String, enum: ['new', 'read', 'resolved'], default: 'new' },
}, { timestamps: true });

contactMessageSchema.index({ status: 1, createdAt: -1 });

const ContactMessage = model('ContactMessage', contactMessageSchema);

// ---- College Interest Enquiry (lead capture before viewing/comparing colleges) ----
const collegeInterestSchema = new Schema({
  name: { type: String, required: true },
  phone: { type: String, required: true },
  email: { type: String, default: '' },
  student: { type: Schema.Types.ObjectId, ref: 'Student' }, // set automatically if logged in
  colleges: [{ type: Schema.Types.ObjectId, ref: 'College' }], // which college(s) they were interested in
  context: { type: String, enum: ['compare', 'view'], default: 'compare' },
  status: { type: String, enum: ['new', 'contacted', 'closed'], default: 'new' },
}, { timestamps: true });

collegeInterestSchema.index({ createdAt: -1 });
collegeInterestSchema.index({ status: 1 });

const CollegeInterest = model('CollegeInterest', collegeInterestSchema);


const reviewSchema = new Schema({
  fullName: { type: String, required: true, trim: true },
  // "Ananya Rao" -> "Ananya R." — computed once on create so the public
  // endpoint never has to hand out a full name.
  displayName: { type: String, default: '' },
  email: { type: String, default: '', lowercase: true, trim: true },
  phone: { type: String, default: '' },
 
  // Free-text context line under the name, e.g. "Engineering · Bangalore"
  course: { type: String, default: '' },
  city: { type: String, default: '' },
 
  // Which product they're talking about — drives the little tag on the card
  source: { type: String, enum: ['KCET', 'PGCET', 'Assessment', 'Counselling', 'Colleges', 'Other'], default: 'Other' },
 
  rating: { type: Number, required: true, min: 1, max: 5 },
  review: { type: String, required: true, trim: true, maxlength: 1000 },
 
  student: { type: Schema.Types.ObjectId, ref: 'Student' }, // set if logged in
 
  // Explicit, recorded permission to publish. Required — no consent, no insert.
  consentToPublish: { type: Boolean, required: true },
  consentAt: { type: Date },
 
  status: { type: String, enum: ['pending', 'approved', 'rejected'], default: 'pending' },
  featured: { type: Boolean, default: false }, // the big pull-quote at the top
  order: { type: Number, default: 0 },
  adminNote: { type: String, default: '' },
  approvedAt: { type: Date },
}, { timestamps: true });
 
reviewSchema.index({ status: 1, featured: -1, order: 1, createdAt: -1 });
reviewSchema.index({ createdAt: -1 });
 
function toDisplayName(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return 'Student';
  if (parts.length === 1) return parts[0];
  return `${parts[0]} ${parts[parts.length - 1][0].toUpperCase()}.`;
}
 
const Review = model('Review', reviewSchema);

/* ============================================================================
 *  AUTH HELPERS (kept inline - not a separate middleware file)
 * ==========================================================================*/
function generateToken(payload) {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: JWT_EXPIRES_IN });
}

function authRequired(req, res, next) {
  try {
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (!token) return res.status(401).json({ success: false, message: 'No token provided' });
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded; // { id, role: 'student' | 'admin' }
    next();
  } catch (err) {
    return res.status(401).json({ success: false, message: 'Invalid or expired token' });
  }
}

function adminOnly(req, res, next) {
  authRequired(req, res, () => {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ success: false, message: 'Admin access only' });
    }
    next();
  });
}

function makeReferralCode(name) {
  const base = (name || 'MMC').replace(/[^a-zA-Z]/g, '').toUpperCase().slice(0, 5) || 'MMC';
  return base + Math.floor(1000 + Math.random() * 9000);
}

function makeInvoiceNumber() {
  return 'INV-' + Date.now() + '-' + Math.floor(100 + Math.random() * 900);
}

// Reused by predictors, assessment, and the dashboard - single source of truth
// for "does this student currently have paid access".
async function getActiveSubscription(studentId) {
  return Subscription.findOne({ student: studentId, status: 'active', endDate: { $gte: new Date() } })
    .populate('plan', 'name price durationInDays')
    .lean();
}


/* ---------------- PROMO HELPERS ---------------- */
// Same rounding rule as the frontend: whole rupees, ₹1 floor unless fully free.
function applyDiscount(price, promo) {
  let discount = promo.discountType === 'percent'
    ? (price * promo.discountValue) / 100
    : promo.discountValue;
  if (promo.maxDiscount > 0) discount = Math.min(discount, promo.maxDiscount);
  const raw = Math.max(0, price - discount);
  if (raw === 0) return 0;
  return Math.max(1, Math.round(raw));
}

// Returns { ok, message, promo, finalPrice, discount }
async function checkPromo(rawCode, plan, studentId) {
  const code = String(rawCode || '').trim().toUpperCase();
  if (!code) return { ok: false, message: 'Enter a promo code first.' };

  const promo = await PromoCode.findOne({ code, isActive: true });
  if (!promo) return { ok: false, message: 'Invalid or expired promo code.' };

  const now = new Date();
  if (promo.validFrom && now < promo.validFrom) return { ok: false, message: 'This promo code is not active yet.' };
  if (promo.validUntil && now > promo.validUntil) return { ok: false, message: 'This promo code has expired.' };
  if (promo.usageLimit > 0 && promo.usedCount >= promo.usageLimit) {
    return { ok: false, message: 'This promo code has been fully redeemed.' };
  }
  if (plan && promo.applicablePlans.length && !promo.applicablePlans.some((id) => String(id) === String(plan._id))) {
    return { ok: false, message: 'This promo code is not valid for the selected plan.' };
  }
  if (studentId && promo.perStudentLimit > 0) {
    const used = await PromoRedemption.countDocuments({ promo: promo._id, student: studentId });
    if (used >= promo.perStudentLimit) return { ok: false, message: 'You have already used this promo code.' };
  }

  const finalPrice = plan ? applyDiscount(plan.price, promo) : null;
  return { ok: true, promo, finalPrice, discount: plan ? plan.price - finalPrice : 0 };
}

// Atomic: only increments if the usage limit still has room.
async function redeemPromo(promoId, studentId, paymentId) {
  const r = await PromoCode.updateOne(
    { _id: promoId, $or: [{ usageLimit: 0 }, { $expr: { $lt: ['$usedCount', '$usageLimit'] } }] },
    { $inc: { usedCount: 1 } }
  );
  if (!r.modifiedCount) return false;
  await PromoRedemption.create({ promo: promoId, student: studentId, payment: paymentId });
  return true;
}

async function activateSubscription(payment, plan) {
  const endDate = new Date(Date.now() + plan.durationInDays * 24 * 60 * 60 * 1000);
  const sub = await Subscription.create({
    student: payment.student, plan: plan._id, endDate, amountPaid: payment.amount, paymentId: payment._id,
  });
  await Referral.updateMany(
    { referredStudent: payment.student, type: 'subscription', status: 'pending' },
    { status: 'converted' }
  );
  return sub;
}




/* ----------------------------------------------------------------------------
 *  PAGINATION HELPER
 *  Every admin "list" endpoint uses this instead of returning the full
 *  collection, so response size stays bounded no matter how much data exists.
 * --------------------------------------------------------------------------*/
function getPagination(req) {
  let page = parseInt(req.query.page, 10);
  let limit = parseInt(req.query.limit, 10);
  if (!Number.isFinite(page) || page < 1) page = 1;
  if (!Number.isFinite(limit) || limit < 1) limit = DEFAULT_PAGE_SIZE;
  if (limit > MAX_PAGE_SIZE) limit = MAX_PAGE_SIZE;
  return { page, limit, skip: (page - 1) * limit };
}

function paginatedResponse(res, { data, total, page, limit, extraKey = 'data' }) {
  return res.json({
    success: true,
    [extraKey]: data,
    pagination: { page, limit, total, totalPages: Math.max(1, Math.ceil(total / limit)) },
  });
}

/* ----------------------------------------------------------------------------
 *  LIGHTWEIGHT INPUT VALIDATION HELPER (no extra dependency)
 *  rules: { fieldName: { required, type, min, max, enum } }
 * --------------------------------------------------------------------------*/
function validateBody(body, rules) {
  const errors = [];
  for (const [field, rule] of Object.entries(rules)) {
    const value = body[field];
    if (rule.required && (value === undefined || value === null || value === '')) {
      errors.push(`${field} is required`);
      continue;
    }
    if (value === undefined || value === null || value === '') continue;
    if (rule.type === 'number' && Number.isNaN(Number(value))) errors.push(`${field} must be a number`);
    if (rule.type === 'email' && !/^\S+@\S+\.\S+$/.test(String(value))) errors.push(`${field} must be a valid email`);
    if (rule.min !== undefined && Number(value) < rule.min) errors.push(`${field} must be >= ${rule.min}`);
    if (rule.enum && !rule.enum.includes(value)) errors.push(`${field} must be one of: ${rule.enum.join(', ')}`);
    if (rule.minLength && String(value).length < rule.minLength) errors.push(`${field} must be at least ${rule.minLength} characters`);
  }
  return errors;
}

function validate(rules) {
  return (req, res, next) => {
    const errors = validateBody(req.body, rules);
    if (errors.length) return res.status(400).json({ success: false, message: 'Validation failed', errors });
    next();
  };
}

/* ----------------------------------------------------------------------------
 *  SIMPLE IN-MEMORY CACHE (per-process)
 *  Good enough for a single-instance deployment. For multi-instance/production
 *  scale, swap this for Redis (same get/set/invalidate call sites).
 *  Used only for public, read-heavy, rarely-changing endpoints (sliders,
 *  colleges, subscription plans, courses, page content).
 * --------------------------------------------------------------------------*/
const cacheStore = new Map();
const CACHE_TTL_MS = Number(process.env.CACHE_TTL_MS || 60 * 1000); // 60s default

function cacheGet(key) {
  const entry = cacheStore.get(key);
  if (!entry) return null;
  if (Date.now() > entry.expiresAt) { cacheStore.delete(key); return null; }
  return entry.value;
}
function cacheSet(key, value, ttl = CACHE_TTL_MS) {
  cacheStore.set(key, { value, expiresAt: Date.now() + ttl });
}
function cacheInvalidate(prefix) {
  for (const key of cacheStore.keys()) {
    if (key.startsWith(prefix)) cacheStore.delete(key);
  }
}
// Wrap a public GET handler with caching in one line: cached('colleges', handler)
function cached(keyPrefix, handler) {
  return async (req, res) => {
    const cacheKey = keyPrefix + ':' + JSON.stringify(req.query);
    const hit = cacheGet(cacheKey);
    if (hit) {
      res.set('X-Cache', 'HIT');
      return res.json(hit);
    }
    res.set('X-Cache', 'MISS');
    const originalJson = res.json.bind(res);
    res.json = (body) => { cacheSet(cacheKey, body); return originalJson(body); };
    return handler(req, res);
  };
}

/* ============================================================================
 *  SEED DEFAULT ADMIN + SUBSCRIPTION PLANS ON STARTUP
 * ==========================================================================*/
async function seedDefaults() {
  try {
    const adminEmail = (process.env.ADMIN_EMAIL || 'admin@mapmycareer360.com').toLowerCase();
    const existingAdmin = await Admin.findOne({ email: adminEmail });
    if (!existingAdmin) {
      const hashed = await bcrypt.hash(process.env.ADMIN_PASSWORD || 'Admin@12345', 10);
      await Admin.create({ name: 'MapMyCareer360 Admin', email: adminEmail, password: hashed, role: 'superadmin' });
      console.log(`[Seed] Default admin created -> ${adminEmail}`);
    }

    const planCount = await SubscriptionPlan.countDocuments();
    if (planCount === 0) {
      // LAUNCH PRICING: all plans temporarily ₹1 while the site is new -
      // change via Admin Panel → Subscriptions → Plans whenever you're ready.
      await SubscriptionPlan.insertMany([
        { name: 'Monthly', durationInDays: 30, price: 1, features: ['Career Assessment access', 'KCET Predictor', 'PGCET Predictor'] },
        { name: 'Quarterly', durationInDays: 90, price: 1, features: ['Career Assessment access', 'KCET Predictor', 'PGCET Predictor', 'Priority support'] },
        { name: 'Yearly', durationInDays: 365, price: 1, features: ['Career Assessment access', 'KCET Predictor', 'PGCET Predictor', 'Priority support', '1-on-1 counselling session'] },
      ]);
      console.log('[Seed] Default subscription plans created (launch pricing ₹1)');
    }

    const sliderCount = await Slider.countDocuments();
    if (sliderCount === 0) {
      await Slider.insertMany([
        { title: 'Map Your Career With Confidence', subtitle: 'Take our scientific Career Assessment and discover the path that fits you.', image: '/images/banner-1.jpg', ctaText: 'Take Free Career Assessment', ctaLink: '/career-assessment', order: 1 },
        { title: 'From Student to Professional', subtitle: 'We guide you every step of the way.', image: '/images/student-to-professional.png', ctaText: 'Explore Subscription Plans', ctaLink: '/subscription', order: 2 },
      ]);
      console.log('[Seed] Default sliders created');
    }

    // Seed the career database with the same 6 fields the frontend already
    // ships as a local fallback, so /api/career-fields isn't empty on launch.
    // Admins can edit/expand these from the Admin Panel at any time.
    const careerFieldCount = await CareerField.countDocuments();
    if (careerFieldCount === 0) {
      await CareerField.insertMany([
        {
          name: 'Engineering & Technology',
          stream: 'Science (PCM) leading into Engineering or Computer Science',
          careers: [
            { name: 'Software Engineer', blurb: 'Builds the applications and systems people use every day.' },
            { name: 'Mechanical Engineer', blurb: 'Designs and improves machines, vehicles, and manufacturing systems.' },
            { name: 'Electronics Engineer', blurb: 'Works on circuits, devices, and embedded systems.' },
            { name: 'Data Engineer', blurb: 'Builds the pipelines that move and organize large datasets.' },
          ],
          courses: ['B.E./B.Tech (CSE, ECE, Mechanical)', 'BCA', 'Diploma in Engineering'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Take PCM; start basic coding or electronics as a hobby.' },
            { stage: 'Entrance exam', detail: 'Prepare for KCET / JEE / COMEDK.' },
            { stage: "Bachelor's degree", detail: 'B.E./B.Tech in a chosen specialization.' },
            { stage: 'Internship', detail: 'Get hands-on project or internship experience in the final year.' },
            { stage: 'First role', detail: 'Join as a junior engineer/developer and specialize over 2-3 years.' },
          ],
          workStyle: 'Structured, problem-solving, often project- or team-based.',
          growth: 'Broad and steadily growing, especially in software and electronics.',
        },
        {
          name: 'Medicine & Healthcare',
          stream: 'Science (PCB) leading into Medicine or Allied Health',
          careers: [
            { name: 'Doctor (MBBS)', blurb: 'Diagnoses and treats patients across general or specialized care.' },
            { name: 'Physiotherapist', blurb: 'Helps patients recover movement and manage physical pain.' },
            { name: 'Medical Lab Technologist', blurb: 'Runs diagnostic tests that guide treatment decisions.' },
            { name: 'Nutritionist', blurb: 'Designs diet and wellness plans for individuals or clinics.' },
          ],
          courses: ['MBBS', 'BPT (Physiotherapy)', 'B.Sc. Nursing', 'B.Sc. Medical Lab Technology'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Take PCB; build strong biology and chemistry fundamentals.' },
            { stage: 'Entrance exam', detail: 'Prepare for NEET or the relevant allied-health entrance test.' },
            { stage: "Bachelor's degree", detail: 'MBBS / BPT / B.Sc. in the chosen specialization.' },
            { stage: 'Clinical training', detail: 'Internship or residency under supervision.' },
            { stage: 'Practice', detail: 'Begin practice, then pursue further specialization if desired.' },
          ],
          workStyle: 'People-focused, high responsibility, often long training periods.',
          growth: 'Consistently in demand, with many specialization paths.',
        },
        {
          name: 'Commerce & Business',
          stream: 'Commerce leading into Business, Finance, or Management',
          careers: [
            { name: 'Chartered Accountant', blurb: 'Manages accounts, audits, and tax compliance for businesses.' },
            { name: 'Business Analyst', blurb: 'Studies data and processes to guide business decisions.' },
            { name: 'Entrepreneur', blurb: 'Builds and runs a business, from idea to operations.' },
            { name: 'Marketing Manager', blurb: 'Plans how a product or brand reaches its customers.' },
          ],
          courses: ['B.Com', 'BBA', 'CA / CS / CMA', 'BBA in Marketing or Finance'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Take Commerce; build comfort with numbers and current affairs.' },
            { stage: 'Entrance/foundation', detail: 'CA Foundation, CUET, or direct BBA/B.Com admission.' },
            { stage: "Bachelor's degree", detail: 'B.Com / BBA, with a chosen elective focus.' },
            { stage: 'Internship', detail: 'Internship in finance, operations, or marketing.' },
            { stage: 'First role', detail: 'Join as an analyst/associate and move into specialization or leadership.' },
          ],
          workStyle: 'Analytical and people-facing, mixes numbers with communication.',
          growth: 'Wide range of roles across every industry.',
        },
        {
          name: 'Arts & Humanities',
          stream: 'Arts / Humanities leading into Law, Civil Services, or Social Sciences',
          careers: [
            { name: 'Lawyer', blurb: 'Advises and represents people on legal matters.' },
            { name: 'Civil Services Officer', blurb: 'Administers public policy and government services.' },
            { name: 'Journalist', blurb: 'Researches and reports on current events.' },
            { name: 'Psychologist', blurb: 'Studies behaviour and supports mental well-being.' },
          ],
          courses: ['BA LLB', 'BA (Political Science / Economics / Psychology)', 'Mass Communication'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Take Arts/Humanities; read widely and follow current affairs.' },
            { stage: 'Entrance exam', detail: 'CLAT (law), CUET, or direct BA admission.' },
            { stage: "Bachelor's degree", detail: 'BA / BA LLB in the chosen specialization.' },
            { stage: 'Practical exposure', detail: 'Internship, moot courts, or fieldwork depending on the path.' },
            { stage: 'First role', detail: 'Join a firm, newsroom, or civil-services preparation track.' },
          ],
          workStyle: 'Reading- and discussion-heavy, values clear communication.',
          growth: 'Steady, with strong long-term paths in law and public service.',
        },
        {
          name: 'Design & Creative Media',
          stream: 'Design or Creative Media foundation programme',
          careers: [
            { name: 'UX/UI Designer', blurb: 'Shapes how digital products look, feel, and work.' },
            { name: 'Graphic Designer', blurb: 'Creates visual content for brands and media.' },
            { name: 'Animator', blurb: 'Brings characters and stories to life on screen.' },
            { name: 'Architect', blurb: 'Designs buildings and spaces people live and work in.' },
          ],
          courses: ['B.Des', 'Bachelor of Fine Arts', 'B.Arch', 'Animation & Multimedia diploma'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Build a portfolio; explore sketching, digital tools, or photography.' },
            { stage: 'Entrance exam', detail: 'NID/NIFT/UCEED or portfolio-based admission.' },
            { stage: "Bachelor's degree", detail: 'B.Des / B.Arch / BFA in the chosen craft.' },
            { stage: 'Studio experience', detail: 'Internship at a studio, agency, or firm.' },
            { stage: 'First role', detail: 'Join as a junior designer and build a specialization.' },
          ],
          workStyle: 'Visual and iterative, mixes independent work with feedback rounds.',
          growth: 'Growing fast in digital product and animation fields.',
        },
        {
          name: 'Science & Research',
          stream: 'Science (PCM/PCB) leading into pure sciences or research',
          careers: [
            { name: 'Research Scientist', blurb: 'Runs experiments to expand knowledge in a chosen field.' },
            { name: 'Data Scientist', blurb: 'Finds patterns in data to answer real-world questions.' },
            { name: 'Environmental Scientist', blurb: 'Studies ecosystems and climate to inform policy.' },
            { name: 'Astrophysicist', blurb: 'Studies the physics of stars, planets, and the universe.' },
          ],
          courses: ['B.Sc. (Physics/Chemistry/Biology/Statistics)', 'Integrated M.Sc.', 'B.S. Research programmes'],
          roadmap: [
            { stage: 'Class 11-12', detail: 'Take Science; build strong fundamentals in the subject of interest.' },
            { stage: 'Entrance exam', detail: 'CUET or institute-specific research-track entrance exams.' },
            { stage: "Bachelor's/Integrated degree", detail: 'B.Sc. or integrated M.Sc. in the chosen subject.' },
            { stage: 'Research exposure', detail: 'Lab projects, papers, or a research internship.' },
            { stage: 'Specialization', detail: 'Pursue M.Sc./Ph.D. or move into applied data roles.' },
          ],
          workStyle: 'Curiosity-driven, methodical, often lab- or data-based.',
          growth: 'Strong in data-driven fields; slower but stable in pure research.',
        },
      ]);
      console.log('[Seed] Default career fields created');
    }

    const careerValueCount = await CareerValue.countDocuments();
    if (careerValueCount === 0) {
      await CareerValue.insertMany([
        { key: 'stability', label: 'Money & Stability', nudge: { 'Commerce & Business': 2, 'Medicine & Healthcare': 1 } },
        { key: 'creativity', label: 'Creativity & Expression', nudge: { 'Design & Creative Media': 2, 'Arts & Humanities': 1 } },
        { key: 'helping', label: 'Helping Others', nudge: { 'Medicine & Healthcare': 2, 'Arts & Humanities': 1 } },
        { key: 'leadership', label: 'Leadership & Influence', nudge: { 'Commerce & Business': 2, 'Arts & Humanities': 1 } },
        { key: 'independence', label: 'Independence & Flexibility', nudge: { 'Design & Creative Media': 1, 'Commerce & Business': 1 } },
        { key: 'learning', label: 'Continuous Learning', nudge: { 'Science & Research': 2, 'Engineering & Technology': 1 } },
        { key: 'structure', label: 'Structure & Routine', nudge: { 'Engineering & Technology': 1, 'Medicine & Healthcare': 1 } },
        { key: 'adventure', label: 'Adventure & Variety', nudge: { 'Arts & Humanities': 1, 'Design & Creative Media': 1 } },
      ]);
      console.log('[Seed] Default career values created');
    }

    // Seed a real starter question bank across all 4 sections so the
    // assessment works out of the box - admins can edit/replace these from
    // the Admin Panel's Career Assessment > Question Bank tab at any time.
    // Weight keys must match CareerField.name exactly.
    const questionCount = await AssessmentQuestion.countDocuments();
    if (questionCount === 0) {
      const ENGG = 'Engineering & Technology', MED = 'Medicine & Healthcare', COM = 'Commerce & Business',
        ARTS = 'Arts & Humanities', DESIGN = 'Design & Creative Media', SCI = 'Science & Research';
      await AssessmentQuestion.insertMany([
        // ---- INTEREST ----
        { category: 'interest', question: 'Which activity would you enjoy most on a free weekend?', options: [
          { label: 'Fixing a broken gadget or building something', weights: { [ENGG]: 3, [SCI]: 1 } },
          { label: 'Reading about diseases, the body, or fitness', weights: { [MED]: 3 } },
          { label: 'Planning how a small business could make money', weights: { [COM]: 3 } },
          { label: 'Sketching, editing videos, or designing something', weights: { [DESIGN]: 3 } },
        ] },
        { category: 'interest', question: 'Which school subject do you look forward to the most?', options: [
          { label: 'Mathematics or Physics', weights: { [ENGG]: 2, [SCI]: 2 } },
          { label: 'Biology', weights: { [MED]: 3 } },
          { label: 'Economics or Business Studies', weights: { [COM]: 3 } },
          { label: 'History, Civics, or Languages', weights: { [ARTS]: 3 } },
        ] },
        { category: 'interest', question: 'Which of these videos would you click on first?', options: [
          { label: '"How rockets actually work"', weights: { [ENGG]: 2, [SCI]: 3 } },
          { label: '"A day in the life of a surgeon"', weights: { [MED]: 3 } },
          { label: '"How this startup grew to ₹100 crore"', weights: { [COM]: 3 } },
          { label: '"Behind the scenes of a movie poster design"', weights: { [DESIGN]: 3 } },
        ] },
        { category: 'interest', question: 'In a general-knowledge quiz, which round would you pick?', options: [
          { label: 'Science & Technology', weights: { [SCI]: 3, [ENGG]: 1 } },
          { label: 'Current Affairs & Politics', weights: { [ARTS]: 3 } },
          { label: 'Business & Economy', weights: { [COM]: 3 } },
          { label: 'Art, Films & Design', weights: { [DESIGN]: 3 } },
        ] },
        { category: 'interest', question: 'Which cause would you most want to volunteer for?', options: [
          { label: 'A free health check-up camp', weights: { [MED]: 3 } },
          { label: 'Teaching coding to school kids', weights: { [ENGG]: 2, [SCI]: 1 } },
          { label: 'A legal-aid or awareness drive', weights: { [ARTS]: 3 } },
          { label: 'Designing posters for a local NGO', weights: { [DESIGN]: 3 } },
        ] },

        // ---- APTITUDE ----
        { category: 'aptitude', question: 'A pattern shows 2, 6, 12, 20, 30 ... what comes next?', options: [
          { label: '42', weights: { [ENGG]: 3, [SCI]: 2 } },
          { label: '40', weights: { [COM]: 1 } },
          { label: '38', weights: {} },
          { label: 'Not sure, I\'d rather estimate', weights: { [ARTS]: 1 } },
        ] },
        { category: 'aptitude', question: 'You\'re given a messy room and 15 minutes. What do you do first?', options: [
          { label: 'Make a quick plan/order before touching anything', weights: { [ENGG]: 2, [COM]: 2 } },
          { label: 'Start grouping similar items by type', weights: { [SCI]: 2, [MED]: 1 } },
          { label: 'Focus on making it look good, not just tidy', weights: { [DESIGN]: 3 } },
          { label: 'Ask what the room is actually needed for', weights: { [ARTS]: 2 } },
        ] },
        { category: 'aptitude', question: 'Which puzzle would you enjoy solving the most?', options: [
          { label: 'A logic/number puzzle (like Sudoku)', weights: { [ENGG]: 2, [SCI]: 2 } },
          { label: 'A word or riddle-based puzzle', weights: { [ARTS]: 3 } },
          { label: 'A visual jigsaw or spot-the-difference', weights: { [DESIGN]: 3 } },
          { label: 'A "who profits from this" business case', weights: { [COM]: 3 } },
        ] },
        { category: 'aptitude', question: 'How comfortable are you memorising detailed diagrams (e.g. the human body)?', options: [
          { label: 'Very - I enjoy detailed diagrams', weights: { [MED]: 3, [SCI]: 1 } },
          { label: 'Fairly, if it\'s a system I can logically map out', weights: { [ENGG]: 2 } },
          { label: 'I prefer big-picture ideas over fine detail', weights: { [ARTS]: 2, [COM]: 1 } },
          { label: 'I\'d rather visualise it as an image/diagram myself', weights: { [DESIGN]: 2 } },
        ] },
        { category: 'aptitude', question: 'You have to explain a complex idea to a 10-year-old. You\'d:', options: [
          { label: 'Break it into clear logical steps', weights: { [ENGG]: 2, [SCI]: 1 } },
          { label: 'Use a story or real-life example', weights: { [ARTS]: 3 } },
          { label: 'Draw a picture or diagram', weights: { [DESIGN]: 3 } },
          { label: 'Relate it to something they\'d "buy" or want', weights: { [COM]: 2 } },
        ] },

        // ---- PERSONALITY ----
        { category: 'personality', question: 'In a group project, you naturally end up:', options: [
          { label: 'Organising tasks and keeping things on track', weights: { [ENGG]: 2, [COM]: 2 } },
          { label: 'Taking care that everyone feels included', weights: { [MED]: 2, [ARTS]: 1 } },
          { label: 'Pitching ideas and leading the direction', weights: { [COM]: 3 } },
          { label: 'Making the final output look/sound great', weights: { [DESIGN]: 3 } },
        ] },
        { category: 'personality', question: 'When something goes wrong, your first instinct is to:', options: [
          { label: 'Diagnose exactly what broke, step by step', weights: { [ENGG]: 3, [SCI]: 1 } },
          { label: 'Check on the people affected first', weights: { [MED]: 3 } },
          { label: 'Think about how to prevent it happening again (a policy/rule)', weights: { [ARTS]: 2 } },
          { label: 'Figure out the cost/impact and next steps', weights: { [COM]: 2 } },
        ] },
        { category: 'personality', question: 'Which environment would you thrive in?', options: [
          { label: 'A quiet lab or workshop, deep-focus work', weights: { [SCI]: 2, [ENGG]: 2 } },
          { label: 'A busy hospital or clinic, always moving', weights: { [MED]: 3 } },
          { label: 'A fast-paced office with targets and deals', weights: { [COM]: 3 } },
          { label: 'A creative studio full of ideas and visuals', weights: { [DESIGN]: 3 } },
        ] },
        { category: 'personality', question: 'How do you usually make decisions?', options: [
          { label: 'Data and evidence first', weights: { [SCI]: 3, [ENGG]: 1 } },
          { label: 'Gut feeling and empathy for people involved', weights: { [MED]: 2, [ARTS]: 1 } },
          { label: 'Risk vs. reward, like a business call', weights: { [COM]: 3 } },
          { label: 'What feels right creatively/aesthetically', weights: { [DESIGN]: 2 } },
        ] },
        { category: 'personality', question: 'Which best describes you under pressure?', options: [
          { label: 'Calm and methodical - I follow a process', weights: { [ENGG]: 2, [MED]: 1 } },
          { label: 'I stay composed to reassure others', weights: { [MED]: 2, [ARTS]: 1 } },
          { label: 'I get energised and push for a quick decision', weights: { [COM]: 2 } },
          { label: 'I need space to think before reacting', weights: { [SCI]: 1, [DESIGN]: 1 } },
        ] },

        // ---- CAREER ORIENTATION (adaptive) ----
        { category: 'adaptive', adaptiveForField: ENGG, question: 'If you lean technical: which excites you more?', options: [
          { label: 'Designing/building physical or digital systems', weights: { [ENGG]: 3 } },
          { label: 'Researching how or why something works', weights: { [SCI]: 3 } },
          { label: 'Making technology easy and beautiful to use', weights: { [DESIGN]: 2, [ENGG]: 1 } },
          { label: 'Selling or managing tech products', weights: { [COM]: 2, [ENGG]: 1 } },
        ] },
        { category: 'adaptive', adaptiveForField: MED, question: 'If you lean people-focused: which role fits best?', options: [
          { label: 'Directly treating/caring for patients', weights: { [MED]: 3 } },
          { label: 'Advocating for people\'s rights or wellbeing', weights: { [ARTS]: 3 } },
          { label: 'Leading a team that serves customers', weights: { [COM]: 2, [MED]: 1 } },
          { label: 'Researching what helps people, at scale', weights: { [SCI]: 2, [MED]: 1 } },
        ] },
        { category: 'adaptive', adaptiveForField: COM, question: 'If you lean business-minded: what\'s more "you"?', options: [
          { label: 'Building and growing your own venture', weights: { [COM]: 3 } },
          { label: 'Making sharp, well-reasoned calls with numbers', weights: { [COM]: 2, [SCI]: 1 } },
          { label: 'Leading people through change', weights: { [ARTS]: 2, [COM]: 1 } },
          { label: 'Designing the brand/product people buy', weights: { [DESIGN]: 2, [COM]: 1 } },
        ] },
        { category: 'adaptive', adaptiveForField: DESIGN, question: 'If you lean creative: which appeals most?', options: [
          { label: 'Designing how digital products look and feel', weights: { [DESIGN]: 3 } },
          { label: 'Telling stories through writing or media', weights: { [ARTS]: 3 } },
          { label: 'Designing physical spaces or products', weights: { [DESIGN]: 2, [ENGG]: 1 } },
          { label: 'Bringing a creative business idea to market', weights: { [COM]: 2, [DESIGN]: 1 } },
        ] },
      ]);
      console.log('[Seed] Default assessment question bank created');
    }
  } catch (err) {
    console.error('[Seed] error:', err.message);
  }
}
mongoose.connection.once('open', seedDefaults);

/* ----------------------------------------------------------------------------
 *  EXCEL CUTOFF-DATA IMPORT HELPERS
 *  Expected columns (header row, case-insensitive, order doesn't matter):
 *    KCET:  Year, College, Course, Category, Round, CutoffRank, 371J
 *    PGCET: Year, College, Course, Category, CollegeType, CutoffRank
 * --------------------------------------------------------------------------*/
const MAX_EXCEL_ROWS = 20000; // hard ceiling per upload to protect the server/DB

/**
 * KEA (Karnataka Examinations Authority) publishes KCET/PGCET cutoff PDFs as
 * one block per college, with course rows whose numbers only stay lined up
 * under their category column when the whitespace layout is preserved -
 * so text is extracted via the `pdftotext -layout` CLI (poppler-utils)
 * rather than a pure-JS PDF library, which loses that column alignment.
 * Requires poppler-utils installed on the server (`apt-get install poppler-utils`).
 */
function extractPdfTextViaPoppler(filePath) {
  try {
    return execFileSync('pdftotext', ['-layout', filePath, '-'], { maxBuffer: 50 * 1024 * 1024 }).toString('utf8');
  } catch (err) {
    if (err.code === 'ENOENT') {
      throw new Error('poppler-utils is not installed on this server. Run: apt-get install -y poppler-utils');
    }
    throw new Error('Could not extract text from PDF: ' + err.message);
  }
}

// Source layout (one block per college):
//   College: E003 B M S College of Engineering, Basavanagudi, Bangalore ...
//   Course Name    1G   1K   1R   2AG   2AK  ...  STR
//   MECHANICAL     13911  --   --  12437   --  ...  --
//   ENGINEERING
// Course names that don't fit one line wrap onto following lines with no
// numbers on them ("ENGINEERING" above) - those get appended to the
// previous data row's name. Every page repeats a footer line and the first
// page has a banner - both are stripped as noise before parsing.
const KEA_PDF_NOISE_PATTERNS = [
  /SC sub category abbreviation/i,
  /^\s*Page \d+ of \d+/i,
  /KARNATAKA EXAMINATIONS AUTHORITY/i,
  /Non-Interactive Admission System/i,
  /ALLOTMENT CUT-OFF RANKS/i,
  /^Seat Type:/i,
];
// Footer fragments that pdftotext glues onto other lines at page breaks
const KEA_GENERATED_RE = /Generated on:\s*[\d-]+\s+[\d:]+/gi;
const KEA_PAGE_RE = /\bPage\s*(\d+\s*)?of(\s*\d+)?/gi;
const KEA_COLLEGE_RE = /^College:\s*(\S+)\s+(.*)$/;
const KEA_HEADER_RE = /^Course Name\s+(.*)$/;
const KEA_VALUE_TOKEN_RE = /^(--|\d+(\.\d+)?)$/;

function cleanKeaCourseName(raw) {
  return raw
    .replace(/Generated on:.*$/i, '')   // safety net: cut any leaked footer/college text
    .replace(/College:.*$/i, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+(\d{1,2})\s+/g, ' ')
    .trim();
}

function parseKeaCutoffPdfText(text) {
  const lines = text
    .split(/\r?\n/)
    .map((ln) => ln.replace(KEA_GENERATED_RE, ' ').replace(KEA_PAGE_RE, ' ').trim()) // trim leading spaces so ^College: matches
    .filter((ln) => ln && !KEA_PDF_NOISE_PATTERNS.some((re) => re.test(ln)));

  const rows = [];
  let currentCollege = null;
  let categories = null;
  let pendingNameParts = [];
  let pendingValues = null;

  const flush = () => {
    if (pendingNameParts.length && pendingValues && currentCollege && categories) {
      const course = cleanKeaCourseName(pendingNameParts.join(' '));
      categories.forEach((cat, i) => {
        const val = pendingValues[i];
        if (val && val !== '--') rows.push({ college: currentCollege, course, category: cat, cutoffrank: Number(val) });
      });
    }
    pendingNameParts = [];
    pendingValues = null;
  };

  for (const raw of lines) {
    const line = raw.trimEnd();
    if (!line.trim()) continue;

    const collegeMatch = KEA_COLLEGE_RE.exec(line);
    if (collegeMatch) {
      flush();
      currentCollege = collegeMatch[2].trim().replace(/\s+/g, ' ');
      categories = null;
      continue;
    }

    const headerMatch = KEA_HEADER_RE.exec(line);
    if (headerMatch) {
      flush();
      categories = headerMatch[1].trim().split(/\s+/);
      continue;
    }

    if (!categories) continue; // stray line before we've seen a header yet

    const tokens = line.trim().split(/\s+/);
    if (tokens.length >= categories.length + 1) {
      const trailing = tokens.slice(-categories.length);
      if (trailing.every((t) => KEA_VALUE_TOKEN_RE.test(t))) {
        flush();
        pendingNameParts = [tokens.slice(0, tokens.length - categories.length).join(' ')];
        pendingValues = trailing;
        continue;
      }
    }

    if (pendingValues) pendingNameParts.push(line.trim()); // wrapped course-name continuation
  }
  flush();

  return rows;
}

function parseExcelRows(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (rows.length > MAX_EXCEL_ROWS) {
      throw new Error(`File has ${rows.length} rows - please split into batches of ${MAX_EXCEL_ROWS} or fewer`);
    }
    // Normalize every row's keys to lowercase/trimmed so "College Name" / "college" / " College " all match
    return rows.map((row) => {
      const normalized = {};
      Object.entries(row).forEach(([key, value]) => {
        normalized[key.trim().toLowerCase().replace(/\s+/g, '')] = typeof value === 'string' ? value.trim() : value;
      });
      return normalized;
    });
  } finally {
    // Uploaded file was only needed transiently for parsing - remove it immediately
    fs.unlink(filePath, () => {});
  }
}

// Resolves college/course names to ObjectIds in bulk (1 query per collection,
// not one per row), auto-creating any college/course names not seen before.
async function resolveCollegeAndCourseIds(rows) {
  const collegeNames = [...new Set(rows.map((r) => String(r.college || '').trim()).filter(Boolean))];
  const courseNames = [...new Set(rows.map((r) => String(r.course || '').trim()).filter(Boolean))];

  const [existingColleges, existingCourses] = await Promise.all([
    College.find({ name: { $in: collegeNames.map((n) => new RegExp(`^${n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i')) } }).select('_id name').lean(),
    Course.find({ name: { $in: courseNames.map((n) => new RegExp(`^${n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i')) } }).select('_id name').lean(),
  ]);

  const collegeMap = new Map(existingColleges.map((c) => [c.name.toLowerCase(), c._id]));
  const courseMap = new Map(existingCourses.map((c) => [c.name.toLowerCase(), c._id]));

  const newColleges = collegeNames.filter((n) => !collegeMap.has(n.toLowerCase())).map((name) => ({ name }));
  const newCourses = courseNames.filter((n) => !courseMap.has(n.toLowerCase())).map((name) => ({ name }));

  if (newColleges.length) {
    const created = await College.insertMany(newColleges, { ordered: false });
    created.forEach((c) => collegeMap.set(c.name.toLowerCase(), c._id));
  }
  if (newCourses.length) {
    const created = await Course.insertMany(newCourses, { ordered: false });
    created.forEach((c) => courseMap.set(c.name.toLowerCase(), c._id));
  }

  return { collegeMap, courseMap };
}

async function bulkUpsertCutoffRows(rows, type) {
  const Model = type === 'kcet' ? KcetCutoff : PgcetCutoff;
  const errors = [];
  const validRows = rows.filter((r, idx) => {
    if (!r.year || !r.college || !r.course || !r.category || !r.cutoffrank) {
      errors.push(`Row ${idx + 2}: missing required field(s)`); // +2 accounts for header row + 0-index
      return false;
    }
    return true;
  });

  if (!validRows.length) return { insertedCount: 0, skipped: rows.length, errors };

  const { collegeMap, courseMap } = await resolveCollegeAndCourseIds(validRows);

  const docs = validRows.map((r) => {
    const base = {
      year: Number(r.year),
      college: collegeMap.get(String(r.college).toLowerCase()),
      course: courseMap.get(String(r.course).toLowerCase()),
      category: String(r.category).toUpperCase(),
      cutoffRank: Number(r.cutoffrank),
    };
    if (type === 'kcet') {
      base.round = r.round || 'Round 1';
    }
    return base;
  });

  // Attach type-specific fields without cluttering the shared mapping above
  validRows.forEach((r, i) => {
    if (type === 'kcet') {
      docs[i].is371J = String(r['371j'] || '').toLowerCase().startsWith('y');
    } else {
      docs[i].collegeType = r.collegetype || 'Private';
    }
  });

  // insertMany with ordered:false lets valid rows succeed even if a few rows fail,
  // and is far faster than looping .create() per row for large sheets.
  const insertResult = await Model.insertMany(docs, { ordered: false }).catch((err) => {
    // Some rows may still have been inserted before the batch error - surface a count if available
    if (err.insertedDocs) return err.insertedDocs;
    errors.push('Some rows failed to insert: ' + err.message);
    return [];
  });

  return {
    insertedCount: Array.isArray(insertResult) ? insertResult.length : 0,
    skipped: rows.length - validRows.length,
    errors,
  };
}

/* ============================================================================
 *  HEALTH CHECK
 * ==========================================================================*/
app.get('/api/health', (req, res) => {
  res.json({ success: true, message: 'MapMyCareer360 API is running', env: NODE_ENV, time: new Date() });
});

/* ============================================================================
 *  STUDENT AUTH: register / login / forgot-password / reset-password / profile
 * ==========================================================================*/
app.post('/api/auth/register', validate({
  fullName: { required: true, minLength: 2 },
  email: { required: true, type: 'email' },
  phone: { required: true, minLength: 10 },
  password: { required: true, minLength: 6 },
}), async (req, res) => {
  try {
    const { fullName, email, phone, password, referredBy } = req.body;
    const existing = await Student.findOne({ email: email.toLowerCase() }).select('_id').lean();
    if (existing) return res.status(409).json({ success: false, message: 'Email already registered' });

    const hashed = await bcrypt.hash(password, 10);
    let referralCode;
    do { referralCode = makeReferralCode(fullName); } while (await Student.findOne({ referralCode }));

    const student = await Student.create({
      fullName, email: email.toLowerCase(), phone, password: hashed, referralCode, referredBy: referredBy || '',
    });

    // If this registration came through a referral link, log it
    if (referredBy) {
      const referrer = await Student.findOne({ referralCode: referredBy });
      if (referrer) {
        await Referral.create({
          referringStudent: referrer._id, referralCode: referredBy,
          referredName: fullName, referredEmail: email, referredStudent: student._id,
          type: 'subscription', status: 'pending',
        });
      }
    }

    const token = generateToken({ id: student._id, role: 'student' });
    res.status(201).json({ success: true, token, student: sanitizeStudent(student) });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/auth/login', validate({
  email: { required: true, type: 'email' },
  password: { required: true },
}), async (req, res) => {
  try {
    const { email, password } = req.body;
    const student = await Student.findOne({ email: (email || '').toLowerCase() });
    if (!student) return res.status(401).json({ success: false, message: 'Invalid credentials' });
    const match = await bcrypt.compare(password, student.password);
    if (!match) return res.status(401).json({ success: false, message: 'Invalid credentials' });
    const token = generateToken({ id: student._id, role: 'student' });
    res.json({ success: true, token, student: sanitizeStudent(student) });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/auth/forgot-password', async (req, res) => {
  try {
    const { email } = req.body;
    const student = await Student.findOne({ email: (email || '').toLowerCase() });
    if (!student) return res.json({ success: true, message: 'If that email exists, a reset link has been generated' });
    const resetToken = jwt.sign({ id: student._id }, JWT_SECRET, { expiresIn: '1h' });
    student.resetToken = resetToken;
    student.resetTokenExpiry = new Date(Date.now() + 60 * 60 * 1000);
    await student.save();
    // NOTE: wire this up to an email/SMS provider in production
    res.json({ success: true, message: 'Reset link generated', resetToken });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/auth/reset-password', async (req, res) => {
  try {
    const { resetToken, newPassword } = req.body;
    const decoded = jwt.verify(resetToken, JWT_SECRET);
    const student = await Student.findOne({ _id: decoded.id, resetToken });
    if (!student || student.resetTokenExpiry < new Date()) {
      return res.status(400).json({ success: false, message: 'Reset link is invalid or expired' });
    }
    student.password = await bcrypt.hash(newPassword, 10);
    student.resetToken = undefined;
    student.resetTokenExpiry = undefined;
    await student.save();
    res.json({ success: true, message: 'Password updated successfully' });
  } catch (err) {
    res.status(400).json({ success: false, message: 'Reset link is invalid or expired' });
  }
});

app.get('/api/students/me', authRequired, async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const student = await Student.findById(req.user.id);
  if (!student) return res.status(404).json({ success: false, message: 'Not found' });
  res.json({ success: true, student: sanitizeStudent(student) });
});

app.put('/api/students/me', authRequired, upload.single('profilePhoto'), async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const allowed = ['fullName', 'phone', 'gender', 'dob', 'address'];
  const updates = {};
  allowed.forEach((k) => { if (req.body[k] !== undefined) updates[k] = req.body[k]; });
  if (req.file) updates.profilePhoto = '/uploads/' + req.file.filename;
  const student = await Student.findByIdAndUpdate(req.user.id, updates, { new: true });
  res.json({ success: true, student: sanitizeStudent(student) });
});

// KYC (bank/UPI details) — required before referral earnings can be paid out.
// Submitting always resets status to "pending" so admin re-reviews any change.
app.put('/api/students/kyc', authRequired, runKycUpload, async (req, res) => {
  const uploaded = Object.values(req.files || {}).flat();
  console.log('[KYC] files received:', Object.keys(req.files || {}), 'dir:', kycDir);
  const discard = () => uploaded.forEach((f) => fs.unlink(f.path, () => {}));
  try {
    if (req.user.role !== 'student') { discard(); return res.status(403).json({ success: false, message: 'Students only' }); }

    const errors = validateBody(req.body, {
      accountHolderName: { required: true, minLength: 2 },
      bankName: { required: true, minLength: 2 },
      accountNumber: { required: true, minLength: 6 },
      ifsc: { required: true, minLength: 6 },
      panNumber: { required: true, minLength: 10 },
    });
    const ifsc = String(req.body.ifsc || '').trim().toUpperCase();
    const pan = String(req.body.panNumber || '').trim().toUpperCase();
    if (!errors.length && !/^[A-Z]{4}0[A-Z0-9]{6}$/.test(ifsc)) errors.push('Enter a valid 11-character IFSC code');
    if (!errors.length && !/^[A-Z]{5}[0-9]{4}[A-Z]$/.test(pan)) errors.push('Enter a valid 10-character PAN');
  if (errors.length) {
  discard();
  console.log('[KYC] rejected:', errors, 'body keys:', Object.keys(req.body));
  return res.status(400).json({ success: false, message: errors.join('. '), errors });
}

    // Every document must exist: either freshly uploaded now, or already on file from before.
    const existing = await Student.findById(req.user.id).select('kyc.documents').lean();
    const oldDocs = existing?.kyc?.documents || {};
    const docs = { passbook: oldDocs.passbook || '', panCard: oldDocs.panCard || '', aadhaarCard: oldDocs.aadhaarCard || '' };
    const replaced = [];
    for (const type of KYC_DOC_TYPES) {
      const file = req.files?.[type]?.[0];
      if (file) {
        if (docs[type]) replaced.push(docs[type]);
        docs[type] = file.filename;
      } else if (!docs[type]) {
        discard();
        return res.status(400).json({ success: false, message: `Please upload your ${KYC_DOC_LABELS[type]}` });
      }
    }

    const student = await Student.findByIdAndUpdate(req.user.id, {
      kyc: {
        accountHolderName: String(req.body.accountHolderName).trim(),
        bankName: String(req.body.bankName).trim(),
        accountNumber: String(req.body.accountNumber).trim(),
        ifsc,
        panNumber: pan,
        upiId: (req.body.upiId || '').trim(),
        documents: docs,
        status: 'pending',
        submittedAt: new Date(),
        verifiedAt: null,
        rejectionReason: '',
      },
    }, { new: true });

    // remove files that were replaced by new uploads
    replaced.forEach((name) => fs.unlink(path.join(kycDir, path.basename(name)), () => {}));
    res.json({ success: true, student: sanitizeStudent(student) });
  } catch (err) {
  discard();
  console.error('[KYC submit failed]', err);
  const status = err.name === 'ValidationError' || err.name === 'CastError' ? 400 : 500;
  res.status(status).json({ success: false, message: err.message });
}
});

function sanitizeStudent(s) {
  const obj = s.toObject ? s.toObject() : s;
  delete obj.password;
  delete obj.resetToken;
  delete obj.resetTokenExpiry;
  if (obj.kyc) {
    const d = obj.kyc.documents || {};
    // the student only needs to know which documents are on file
    obj.kyc.documents = { passbook: !!d.passbook, panCard: !!d.panCard, aadhaarCard: !!d.aadhaarCard };
    delete obj.kyc.reviewedBy;
  }
  return obj;
}

/* ============================================================================
 *  STUDENT DASHBOARD (aggregated: profile + subscription + assessments + payments)
 * ==========================================================================*/
app.get('/api/students/dashboard', authRequired, async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const studentId = req.user.id;
  const [student, subscription, assessments, payments, predictorLeads, collegeInterests] = await Promise.all([
    Student.findById(studentId).lean(),
    Subscription.findOne({ student: studentId, status: 'active', endDate: { $gte: new Date() } }).populate('plan').lean(),
    // Dashboard only needs recent history at a glance - full history lives behind
    // the paginated "Assessments"/"Payments" tabs if we add pagination there later.
    AssessmentResult.find({ student: studentId }).select('-answers').sort({ createdAt: -1 }).limit(25).lean(),
    Payment.find({ student: studentId }).sort({ createdAt: -1 }).limit(25).lean(),
    // examType + accessType tell us, per exam, whether this student has ever
    // unlocked it (via subscription or a one-time payment) - used to render
    // the "Career Tools" access cards without changing how gating works.
    PredictorLead.find({ student: studentId }).select('examType accessType createdAt').sort({ createdAt: -1 }).lean(),
    CollegeInterest.find({ student: studentId }).populate('colleges', 'name location image logo ranking rating').sort({ createdAt: -1 }).limit(12).lean(),
  ]);

  const subscriptionActive = !!subscription;
  const assessmentUnlocked = subscriptionActive || payments.some((p) => p.purpose === 'assessment' && p.status === 'success');
  const kcetUnlocked = subscriptionActive || predictorLeads.some((p) => p.examType === 'kcet');
  const pgcetUnlocked = subscriptionActive || predictorLeads.some((p) => p.examType === 'pgcet');

  // Flatten + de-duplicate the colleges pulled from this student's compare/view interest history.
  const seen = new Set();
  const myColleges = [];
  collegeInterests.forEach((ci) => {
    (ci.colleges || []).forEach((c) => {
      const id = String(c._id);
      if (!c || seen.has(id)) return;
      seen.add(id);
      myColleges.push(c);
    });
  });

  res.json({
    success: true,
    profile: sanitizeStudent(student),
    activeSubscription: subscription,
    assessmentHistory: assessments,
    paymentHistory: payments,
    moduleAccess: {
      subscription: subscriptionActive,
      assessment: assessmentUnlocked,
      kcet: kcetUnlocked,
      pgcet: pgcetUnlocked,
    },
    myColleges: myColleges.slice(0, 8),
  });
});

/* ============================================================================
 *  SUBSCRIPTION PLANS + PURCHASE
 * ==========================================================================*/
app.get('/api/subscription-plans', cached('subscription-plans', async (req, res) => {
  const plans = await SubscriptionPlan.find({ isActive: true }).sort({ price: 1 }).lean();
  res.set('Cache-Control', 'public, max-age=300');
  res.json({ success: true, plans });
}));

// Lightweight check used by predictor pages to decide whether to show
// results or a "subscribe to unlock" paywall.
app.get('/api/subscriptions/status', authRequired, async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const active = await getActiveSubscription(req.user.id);
  res.json({ success: true, active: !!active, subscription: active || null });
});

app.post('/api/subscriptions/purchase', authRequired, validate({ planId: { required: true } }), async (req, res) => {
  try {
    if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
    const { planId, referralCode, promoCode } = req.body;
    const plan = await SubscriptionPlan.findById(planId).lean();
    if (!plan) return res.status(404).json({ success: false, message: 'Plan not found' });

    let promo = null;
    let finalPrice = plan.price;
    if (promoCode) {
      const check = await checkPromo(promoCode, plan, req.user.id);
      if (!check.ok) return res.status(400).json({ success: false, message: check.message });
      promo = check.promo;
      finalPrice = check.finalPrice;
    }

    const promoFields = promo
      ? { promo: promo._id, promoCode: promo.code, originalAmount: plan.price, discountAmount: plan.price - finalPrice }
      : {};

    // ---- 100% discount: no gateway, activate immediately ----
    if (finalPrice === 0) {
      const payment = await Payment.create({
        student: req.user.id, purpose: 'subscription', plan: plan._id, amount: 0,
        gateway: 'razorpay', gatewayPaymentId: 'PROMO-FREE',
        status: 'success', invoiceNumber: makeInvoiceNumber(), ...promoFields,
      });
      const redeemed = await redeemPromo(promo._id, req.user.id, payment._id);
      if (!redeemed) {
        await Payment.findByIdAndDelete(payment._id);
        return res.status(400).json({ success: false, message: 'This promo code has been fully redeemed.' });
      }
      const subscription = await activateSubscription(payment, plan);
      return res.json({ success: true, free: true, payment, plan, subscription, amount: 0 });
    }

    // ---- Normal paid flow ----
    const rzpOrder = await razorpay.orders.create({
      amount: finalPrice * 100, // paise
      currency: 'INR',
      receipt: makeInvoiceNumber(),
    });

    const payment = await Payment.create({
      student: req.user.id, purpose: 'subscription', plan: plan._id, amount: finalPrice,
      gateway: 'razorpay', gatewayOrderId: rzpOrder.id,
      status: 'created', invoiceNumber: rzpOrder.receipt, ...promoFields,
    });

    res.json({
      success: true, payment, plan,
      razorpayOrderId: rzpOrder.id,
      razorpayKeyId: process.env.RAZORPAY_KEY_ID,
      amount: rzpOrder.amount,
      currency: rzpOrder.currency,
    });

    if (referralCode) {
      const referrer = await Student.findOne({ referralCode });
      if (referrer) {
        await Referral.create({
          referringStudent: referrer._id, referralCode, referredStudent: req.user.id,
          type: 'subscription', status: 'pending',
        });
      }
    }
  } catch (err) {
    console.error('[Razorpay order.create failed - subscriptions/purchase]', err.error || err);
    res.status(500).json({
      success: false,
      message: err.error?.description || err.message || 'Payment gateway error',
    });
  }
});
app.post('/api/promo-codes/validate', async (req, res) => {
  try {
    let studentId;
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (token) {
      try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* guest */ }
    }

    const check = await checkPromo(req.body.code, null, studentId);
    if (!check.ok) return res.json({ success: true, valid: false, message: check.message });

    const p = check.promo;
    res.json({
      success: true,
      valid: true,
      code: p.code,
      discountType: p.discountType,
      discountValue: p.discountValue,
      maxDiscount: p.maxDiscount || 0,
      applicablePlans: p.applicablePlans.map(String),
      message: `Promo code ${p.code} applied.`,
    });
  } catch (err) {
    res.status(500).json({ success: false, valid: false, message: 'Could not validate promo code' });
  }
});




// Confirm payment (called after gateway success callback) -> activates subscription
app.post('/api/payments/verify', async (req, res) => {
  try {
    const { paymentId, razorpay_order_id, razorpay_payment_id, razorpay_signature } = req.body;

    const payment = await Payment.findById(paymentId);
    if (!payment) return res.status(404).json({ success: false, message: 'Payment not found' });
    if (payment.gatewayOrderId !== razorpay_order_id) {
      return res.status(400).json({ success: false, message: 'Payment verification failed' });
    }

    const expectedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_KEY_SECRET)
      .update(razorpay_order_id + '|' + razorpay_payment_id)
      .digest('hex');
    if (expectedSignature !== razorpay_signature) {
      return res.status(400).json({ success: false, message: 'Payment verification failed' });
    }

    // Idempotent: a repeated call must not create a second subscription
    if (payment.status === 'success') return res.json({ success: true, payment });

    payment.status = 'success';
    payment.gatewayPaymentId = razorpay_payment_id;
    await payment.save();

    if (payment.promo) {
      const ok = await redeemPromo(payment.promo, payment.student, payment._id);
      if (!ok) console.warn('[Promo] limit reached during checkout race for payment', String(payment._id));
      // customer already paid, so the subscription is still activated
    }

    if (payment.purpose === 'subscription' && payment.plan) {
      const plan = await SubscriptionPlan.findById(payment.plan);
      const sub = await activateSubscription(payment, plan);
      return res.json({ success: true, payment, subscription: sub });
    }

    res.json({ success: true, payment });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
}); 


/* ---- Manage Promo Codes ---- */
app.get('/api/admin/promo-codes', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const [codes, total] = await Promise.all([
    PromoCode.find().populate('applicablePlans', 'name').sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    PromoCode.countDocuments(),
  ]);
  paginatedResponse(res, { data: codes, total, page, limit, extraKey: 'promoCodes' });
});

app.post('/api/admin/promo-codes', adminOnly, validate({
  code: { required: true, minLength: 3 },
  discountType: { required: true, enum: ['percent', 'flat'] },
  discountValue: { required: true, type: 'number', min: 0 },
}), async (req, res) => {
  try {
    const body = { ...req.body, code: String(req.body.code).trim().toUpperCase() };
    if (body.discountType === 'percent' && Number(body.discountValue) > 100) {
      return res.status(400).json({ success: false, message: 'Percent discount cannot exceed 100' });
    }
    const promo = await PromoCode.create(body);
    res.status(201).json({ success: true, promo });
  } catch (err) {
    const dup = err.code === 11000;
    res.status(dup ? 409 : 400).json({ success: false, message: dup ? 'That code already exists' : err.message });
  }
});

app.put('/api/admin/promo-codes/:id', adminOnly, async (req, res) => {
  const updates = { ...req.body };
  delete updates.usedCount;
  if (updates.code) updates.code = String(updates.code).trim().toUpperCase();
  const promo = await PromoCode.findByIdAndUpdate(req.params.id, updates, { new: true, runValidators: true });
  res.json({ success: true, promo });
});

app.delete('/api/admin/promo-codes/:id', adminOnly, async (req, res) => {
  await PromoCode.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'Promo code deleted' });
});

/* ============================================================================
 *  CAREER ASSESSMENT
 * ==========================================================================*/
app.get('/api/assessment/questions', authRequired, async (req, res) => {
  // Grouped by category so the frontend wizard can use these directly wherever
  // it currently falls back to its own local sample question banks.
  const docs = await AssessmentQuestion.find({ isActive: true }).lean();
  const shape = (q) => ({
    id: String(q._id),
    question: q.question,
    options: (q.options || []).map((o) => ({ label: o.label, weights: o.weights || {} })),
  });
  const interest = docs.filter((q) => q.category === 'interest').map(shape);
  const aptitude = docs.filter((q) => q.category === 'aptitude').map(shape);
  const personality = docs.filter((q) => q.category === 'personality').map(shape);
  const adaptive = docs.filter((q) => q.category === 'adaptive')
    .map((q) => ({ ...shape(q), adaptiveForField: q.adaptiveForField || '' }));
  res.json({ success: true, interest, aptitude, personality, adaptive });
});

/* ----------------------------------------------------------------------------
 *  CAREER DATABASE - public read endpoints (fields + values)
 *  Backs the "Career Database" advanced-plan feature. Admins manage these via
 *  /api/admin/career-fields and /api/admin/career-values further down; the
 *  frontend's local CAREER_DB/VALUES objects are just a fallback until these
 *  are populated (they're seeded with matching defaults on first launch).
 * --------------------------------------------------------------------------*/
app.get('/api/career-fields', cached('career-fields', async (req, res) => {
  const fields = await CareerField.find({ isActive: true }).lean();
  res.set('Cache-Control', 'public, max-age=120');
  res.json({ success: true, fields });
}));

app.get('/api/career-values', cached('career-values', async (req, res) => {
  const values = await CareerValue.find({ isActive: true }).lean();
  res.set('Cache-Control', 'public, max-age=120');
  res.json({ success: true, values });
}));

// Start assessment: checks if student has an active subscription (free access) or must pay first
app.get('/api/assessment/access', authRequired, async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const activeSub = await Subscription.findOne({ student: req.user.id, status: 'active', endDate: { $gte: new Date() } });
  res.json({ success: true, freeAccess: !!activeSub, assessmentFee: 199 });
});

app.post('/api/assessment/pay', authRequired, async (req, res) => {
  try {
    const rzpOrder = await razorpay.orders.create({
      amount: 199 * 100,
      currency: 'INR',
      receipt: makeInvoiceNumber(),
    });
    const payment = await Payment.create({
      student: req.user.id, purpose: 'assessment', amount: 199, gateway: 'razorpay',
      gatewayOrderId: rzpOrder.id, status: 'created', invoiceNumber: rzpOrder.receipt,
    });
    res.json({
      success: true, payment,
      razorpayOrderId: rzpOrder.id,
      razorpayKeyId: process.env.RAZORPAY_KEY_ID,
      amount: rzpOrder.amount,
      currency: rzpOrder.currency,
    });
  } catch (err) {
    // Razorpay SDK errors are nested - log the real cause, not just err.message
    console.error('[Razorpay order.create failed]', err.error || err);
    res.status(500).json({
      success: false,
      message: err.error?.description || err.message || 'Payment gateway error',
    });
  }
});

// Friendly "archetype" label from the top 1-2 scoring fields - purely cosmetic,
// makes the report feel personal rather than just a list of scores.
const ARCHETYPE_MAP = {
  'Engineering & Technology': 'The Builder', 'Medicine & Healthcare': 'The Healer',
  'Commerce & Business': 'The Strategist', 'Arts & Humanities': 'The Communicator',
  'Design & Creative Media': 'The Creator', 'Science & Research': 'The Explorer',
};
function archetypeFor(topField, secondField) {
  const top = ARCHETYPE_MAP[topField] || 'The All-Rounder';
  if (!secondField || secondField === topField) return top;
  const secondWord = (ARCHETYPE_MAP[secondField] || '').replace('The ', '');
  return secondWord ? `${top} with a touch of ${secondWord}` : top;
}

app.post('/api/assessment/submit', authRequired, async (req, res) => {
  try {
    if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
    const { studentInfo, academicProfile, careerValues, answers, paymentId, isFreeViaSubscription } = req.body;
    // answers: [{ questionId, selectedOptionIndex }] across all 4 sections combined

    if (!isFreeViaSubscription) {
      const payment = await Payment.findById(paymentId);
      if (!payment || payment.status !== 'success') {
        return res.status(402).json({ success: false, message: 'Payment required before submitting assessment' });
      }
    }

    const questionIds = (answers || []).map((a) => a.questionId).filter((id) => mongoose.isValidObjectId(id));
    const [questions, allValues] = await Promise.all([
      AssessmentQuestion.find({ _id: { $in: questionIds } }).lean(),
      CareerValue.find({ key: { $in: careerValues || [] } }).lean(),
    ]);

    // 1) Tally weighted scores per career field, and per section (for strengths/growth text)
    const scores = {};
    const sectionTotals = { interest: 0, aptitude: 0, personality: 0, adaptive: 0 };
    (answers || []).forEach((a) => {
      const q = questions.find((qq) => String(qq._id) === String(a.questionId));
      const opt = q?.options?.[a.selectedOptionIndex];
      if (!q || !opt) return;
      sectionTotals[q.category] = (sectionTotals[q.category] || 0) + 1;
      Object.entries(opt.weights || {}).forEach(([field, w]) => {
        scores[field] = (scores[field] || 0) + Number(w || 0);
      });
    });
    allValues.forEach((v) => {
      Object.entries(v.nudge || {}).forEach(([field, w]) => {
        scores[field] = (scores[field] || 0) + Number(w || 0);
      });
    });

    const ranked = Object.entries(scores).sort((a, b) => b[1] - a[1]);
    const topField = ranked[0]?.[0] || '';
    const secondField = ranked[1]?.[0] || '';
    const recommendedStreams = ranked.slice(0, 3).map(([field]) => field);

    // 2) Pull the matching CareerField docs for careers/courses/roadmap
    const [topDoc, secondDoc] = await Promise.all([
      topField ? CareerField.findOne({ name: topField, isActive: true }).lean() : null,
      secondField ? CareerField.findOne({ name: secondField, isActive: true }).lean() : null,
    ]);

    const top10Careers = [
      ...(topDoc?.careers || []),
      ...(secondDoc?.careers || []),
    ].slice(0, 10);
    const top3Careers = top10Careers.slice(0, 3);

    const strengths = [];
    const growthAreas = [];
    if (sectionTotals.aptitude >= 3) strengths.push('Strong logical & analytical reasoning');
    else growthAreas.push('Practice structured problem-solving to sharpen analytical thinking');
    if (sectionTotals.interest >= 3) strengths.push('Clear, well-formed interests to build a career around');
    if (sectionTotals.personality >= 3) strengths.push('Good self-awareness of working style and preferences');
    if (!topDoc) growthAreas.push('Explore a wider range of subjects to discover stronger interest signals');
    if (strengths.length === 0) strengths.push('A balanced, adaptable profile across multiple fields');

    const result = {
      archetype: archetypeFor(topField, secondField),
      topField,
      secondField,
      top10Careers,
      top3Careers,
      stream: topDoc?.stream || '',
      courses: topDoc?.courses || [],
      roadmap: topDoc?.roadmap || [],
      workStyle: topDoc?.workStyle || '',
      growth: topDoc?.growth || '',
      strengths,
      growthAreas,
    };

    const created = await AssessmentResult.create({
      student: req.user.id,
      studentInfo: studentInfo || {},
      academicProfile: academicProfile || {},
      careerValues: careerValues || [],
      scores,
      result,
      recommendedStreams,
      isFreeViaSubscription: !!isFreeViaSubscription,
      paymentId: paymentId || undefined,
    });

    res.status(201).json({ success: true, result: created });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/assessment/results/:id', authRequired, async (req, res) => {
  const result = await AssessmentResult.findById(req.params.id);
  if (!result) return res.status(404).json({ success: false, message: 'Not found' });
  res.json({ success: true, result });
});

/* ============================================================================
 *  KCET PREDICTOR
 *  Gated: student must be logged in AND have an active subscription before
 *  results are returned - predictor form can be filled in freely, but
 *  clicking "Predict" for a logged-out/unsubscribed student returns 402 with
 *  a clear next step instead of leaking real cutoff data for free.
 * ==========================================================================*/

// Drives the predictor form's category/course/year dropdowns from whatever
// data has actually been uploaded - so the form can never again offer a
// category (like the old hardcoded "2A"/"SC"/"ST") that doesn't exist in
// the real KEA data and silently returns zero results.
app.get('/api/kcet-cutoffs/meta', cached('kcet-meta', async (req, res) => {
  const [categories, years, courseIds] = await Promise.all([
    KcetCutoff.distinct('category'),
    KcetCutoff.distinct('year'),
    KcetCutoff.distinct('course'),
  ]);
const courses = await Course.find({
  _id: { $in: courseIds },
  name: { $not: /Generated on|College:|Page\s+of/i },
}).select('name').sort({ name: 1 }).lean();

// de-duplicate identical names
const seen = new Set();
const uniqueCourses = courses.filter((c) => {
  const k = c.name.toLowerCase();
  if (seen.has(k)) return false;
  seen.add(k);
  return true;
});

res.json({
  success: true,
  categories: categories.sort(),
  years: years.sort((a, b) => b - a),
  courses: uniqueCourses.map((c) => ({ id: c._id, name: c.name })),
});
}));

app.post('/api/predictors/kcet', validate({
  name: { required: true, minLength: 2 },
  phone: { required: true, minLength: 8 },
  rank: { required: true, type: 'number', min: 1 },
  category: { required: true },
}), async (req, res) => {
  try {
    const { rank, category, is371J, course, year, name, phone, email, paymentId } = req.body;

    // optional auth - logged-in students get checked for an active subscription
    let studentId;
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (token) {
      try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* not logged in, still allow one-time */ }
    }

    let accessType = null;
    if (studentId) {
      const activeSub = await getActiveSubscription(studentId);
      if (activeSub) accessType = 'subscription';
    }
    if (!accessType && ALLOW_DEV_SKIP_PAYMENT && req.body.devSkipPayment) {
      accessType = 'dev-skip';
    }
    if (!accessType) {
      if (!paymentId) {
        return res.status(402).json({
          success: false,
          subscriptionRequired: true,
          oneTimeFee: PREDICTOR_ONE_TIME_FEE,
          message: 'Subscribe or pay a one-time fee to unlock KCET college predictions',
        });
      }
      const payment = await Payment.findById(paymentId);
      if (!payment || payment.status !== 'success' || payment.purpose !== 'predictor') {
        return res.status(402).json({ success: false, message: 'Payment not verified' });
      }
      accessType = 'one-time';
    }

    const filter = { category, cutoffRank: { $gte: Number(rank) } };
    if (year) filter.year = Number(year);
    if (course) filter.course = course;
    if (is371J) filter.is371J = true;

    const matches = await KcetCutoff.find(filter)
      .select('year category round cutoffRank is371J college course')
      .populate('college', 'name location type image ranking')
      .populate('course', 'name level durationYears')
      .sort({ cutoffRank: 1 })
      .limit(MAX_PREDICTOR_RESULTS)
      .lean();

    const withZone = matches.map((m) => {
      const diff = m.cutoffRank - Number(rank);
      let zone = 'Dream';
      if (diff > 3000) zone = 'Safe';
      else if (diff > 500) zone = 'Moderate';
      return { ...m, zone };
    });

    // always log the enquiry, regardless of access path
    await PredictorLead.create({
      name, phone, email: email || '', examType: 'kcet', rank, category, course: course || '',
      student: studentId, accessType, paymentId: accessType === 'one-time' ? paymentId : undefined,
      resultsShown: withZone.length,
    });

    res.json({
      success: true,
      inputRank: rank,
      safe: withZone.filter((m) => m.zone === 'Safe'),
      moderate: withZone.filter((m) => m.zone === 'Moderate'),
      dream: withZone.filter((m) => m.zone === 'Dream'),
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

/* ============================================================================
 *  PGCET PREDICTOR
 *  Same subscription gate as KCET above.
 * ==========================================================================*/
// app.post('/api/predictors/pgcet', authRequired, validate({
//   rank: { required: true, type: 'number', min: 1 },
//   category: { required: true },
// }), async (req, res) => {
//   try {
//     if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
//     const activeSub = await getActiveSubscription(req.user.id);
//     if (!activeSub) {
//       return res.status(402).json({
//         success: false,
//         subscriptionRequired: true,
//         message: 'Subscribe to unlock PGCET college predictions',
//       });
//     }

//     const { rank, category, course, collegeType, year } = req.body;

//     const filter = { category, cutoffRank: { $gte: Number(rank) } };
//     if (year) filter.year = Number(year);
//     if (course) filter.course = course;
//     if (collegeType) filter.collegeType = collegeType;

//     const matches = await PgcetCutoff.find(filter)
//       .select('year category collegeType cutoffRank college course')
//       .populate('college', 'name location type image ranking')
//       .populate('course', 'name level durationYears')
//       .sort({ cutoffRank: 1 })
//       .limit(MAX_PREDICTOR_RESULTS)
//       .lean();

//     const withZone = matches.map((m) => {
//       const diff = m.cutoffRank - Number(rank);
//       let zone = 'Dream';
//       if (diff > 2000) zone = 'Safe';
//       else if (diff > 300) zone = 'Moderate';
//       return { ...m, zone };
//     });

//     res.json({
//       success: true,
//       inputRank: rank,
//       safe: withZone.filter((m) => m.zone === 'Safe'),
//       moderate: withZone.filter((m) => m.zone === 'Moderate'),
//       dream: withZone.filter((m) => m.zone === 'Dream'),
//     });
//   } catch (err) {
//     res.status(500).json({ success: false, message: err.message });
//   }
// });


// Same "drive the form from real data" meta endpoint as KCET above.
app.get('/api/pgcet-cutoffs/meta', cached('pgcet-meta', async (req, res) => {
  const [categories, years, courseIds] = await Promise.all([
    PgcetCutoff.distinct('category'),
    PgcetCutoff.distinct('year'),
    PgcetCutoff.distinct('course'),
  ]);
  const courses = await Course.find({ _id: { $in: courseIds } }).select('name').sort({ name: 1 }).lean();
  res.set('Cache-Control', 'public, max-age=300');
  res.json({
    success: true,
    categories: categories.sort(),
    years: years.sort((a, b) => b - a),
    courses: courses.map((c) => ({ id: c._id, name: c.name })),
  });
}));

app.post('/api/predictors/pgcet', validate({
  name: { required: true, minLength: 2 },
  phone: { required: true, minLength: 8 },
  rank: { required: true, type: 'number', min: 1 },
  category: { required: true },
}), async (req, res) => {
  try {
    const { rank, category, is371J, course, year, collegeType, name, phone, email, paymentId } = req.body;

    // optional auth - logged-in students get checked for an active subscription
    let studentId;
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (token) {
      try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* not logged in, still allow one-time */ }
    }

    let accessType = null;
    if (studentId) {
      const activeSub = await getActiveSubscription(studentId);
      if (activeSub) accessType = 'subscription';
    }
    if (!accessType && ALLOW_DEV_SKIP_PAYMENT && req.body.devSkipPayment) {
      accessType = 'dev-skip';
    }
    if (!accessType) {
      if (!paymentId) {
        return res.status(402).json({
          success: false,
          subscriptionRequired: true,
          oneTimeFee: PREDICTOR_ONE_TIME_FEE,
          message: 'Subscribe or pay a one-time fee to unlock PGCET college predictions',
        });
      }
      const payment = await Payment.findById(paymentId);
      if (!payment || payment.status !== 'success' || payment.purpose !== 'predictor') {
        return res.status(402).json({ success: false, message: 'Payment not verified' });
      }
      accessType = 'one-time';
    }

    const filter = { category, cutoffRank: { $gte: Number(rank) } };
    if (year) filter.year = Number(year);
    if (course) filter.course = course;
    if (is371J) filter.is371J = true;
    if (collegeType) filter.collegeType = collegeType;

    const matches = await PgcetCutoff.find(filter)
      .select('year category round cutoffRank is371J college course')
      .populate('college', 'name location type image ranking')
      .populate('course', 'name level durationYears')
      .sort({ cutoffRank: 1 })
      .limit(MAX_PREDICTOR_RESULTS)
      .lean();

    const withZone = matches.map((m) => {
      const diff = m.cutoffRank - Number(rank);
      let zone = 'Dream';
      if (diff > 3000) zone = 'Safe';
      else if (diff > 500) zone = 'Moderate';
      return { ...m, zone };
    });

    // always log the enquiry, regardless of access path
    await PredictorLead.create({
      name, phone, email: email || '', examType: 'pgcet', rank, category, course: course || '',
      student: studentId, accessType, paymentId: accessType === 'one-time' ? paymentId : undefined,
      resultsShown: withZone.length,
    });

    res.json({
      success: true,
      inputRank: rank,
      safe: withZone.filter((m) => m.zone === 'Safe'),
      moderate: withZone.filter((m) => m.zone === 'Moderate'),
      dream: withZone.filter((m) => m.zone === 'Dream'),
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/predictors/pay', async (req, res) => {
  try {
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    let studentId;
    if (token) {
      try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* guest checkout, fine */ }
    }

    const rzpOrder = await razorpay.orders.create({
      amount: PREDICTOR_ONE_TIME_FEE * 100,
      currency: 'INR',
      receipt: makeInvoiceNumber(),
    });

    const payment = await Payment.create({
      student: studentId, purpose: 'predictor', amount: PREDICTOR_ONE_TIME_FEE,
      gateway: 'razorpay', gatewayOrderId: rzpOrder.id, status: 'created', invoiceNumber: rzpOrder.receipt,
    });

    res.json({
      success: true, payment,
      razorpayOrderId: rzpOrder.id,
      razorpayKeyId: process.env.RAZORPAY_KEY_ID,
      amount: rzpOrder.amount,
      currency: rzpOrder.currency,
    });
  } catch (err) {
    // Razorpay SDK errors are nested - log the real cause, not just err.message
    console.error('[Razorpay order.create failed - predictors/pay]', err.error || err);
    res.status(500).json({
      success: false,
      message: err.error?.description || err.message || 'Payment gateway error',
    });
  }
});






/* ============================================================================
 *  COLLEGES / COURSES / COLLEGE COMPARE (public)
 * ==========================================================================*/
// Paginated + cached: colleges list can grow large, so never returned in full.
// Supports the filters the college hub page needs: text search, type,
// location, specialization and a minimum rating - all optional.
app.get('/api/colleges', cached('colleges', async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.search) filter.name = { $regex: String(req.query.search).slice(0, 60), $options: 'i' };
  if (req.query.type) filter.type = req.query.type;
  if (req.query.location) filter.location = req.query.location;
  if (req.query.specialization) filter.specializations = { $regex: String(req.query.specialization).slice(0, 60), $options: 'i' };
  if (req.query.minRating) filter.rating = { $gte: Number(req.query.minRating) || 0 };

  const sortMap = {
    ranking: { ranking: 1 },
    'fees-asc': { 'fees.tuitionAnnual': 1 },
    'fees-desc': { 'fees.tuitionAnnual': -1 },
    rating: { rating: -1 },
    placement: { 'placements.highestPackage': -1 },
  };
  const sort = sortMap[req.query.sort] || { featured: -1, ranking: 1 };

  const [colleges, total] = await Promise.all([
    College.find(filter)
      .populate('coursesOffered', 'name level')
      .sort(sort)
      .skip(skip).limit(limit)
      .lean(),
    College.countDocuments(filter),
  ]);
  res.set('Cache-Control', 'public, max-age=60');
  paginatedResponse(res, { data: colleges, total, page, limit, extraKey: 'colleges' });
}));

// Lightweight endpoint that hands the hub page the distinct filter-chip
// values (locations, specializations) so it doesn't need to page through
// every college just to build its filter bar.
app.get('/api/colleges/meta/filters', cached('college-filters', async (req, res) => {
  const [locations, specializations, types] = await Promise.all([
    College.distinct('location'),
    College.distinct('specializations'),
    College.distinct('type'),
  ]);
  res.set('Cache-Control', 'public, max-age=300');
  res.json({
    success: true,
    locations: locations.filter(Boolean).sort(),
    specializations: specializations.filter(Boolean).sort(),
    types: types.filter(Boolean),
  });
}));

app.get('/api/colleges/:id', async (req, res) => {
  const college = await College.findById(req.params.id).populate('coursesOffered').lean();
  if (!college) return res.status(404).json({ success: false, message: 'Not found' });
  res.json({ success: true, college });
});

// Compare is capped to 4 colleges on the frontend, but enforce it server-side too.
// Highlights are generated in positive-only language for every college shown -
// we never say one college is "worse", only what each one is good for.
app.post('/api/colleges/compare', async (req, res) => {
  const collegeIds = Array.isArray(req.body.collegeIds) ? req.body.collegeIds.slice(0, 4) : [];
  if (!collegeIds.length) return res.status(400).json({ success: false, message: 'collegeIds is required' });
  const colleges = await College.find({ _id: { $in: collegeIds } }).populate('coursesOffered', 'name level').lean();

  const withHighlights = colleges.map((c) => {
    const highlights = [];
    if (c.ranking && c.ranking <= 20) highlights.push('🏆 Top-Ranked Institution');
    if (c.rating && c.rating >= 4.2) highlights.push('⭐ Highly Rated by Students');
    if (c.placements?.highestPackage >= 10) highlights.push('💼 Strong Placement Record');
    if ((c.fees?.tuitionAnnual || c.fees?.annual) && (c.fees.tuitionAnnual || c.fees.annual) <= 100000) highlights.push('💰 Great Value for Money');
    if ((c.accreditations || []).length >= 2) highlights.push('🎖️ Well Accredited');
    if ((c.coursesOffered || []).length >= 3) highlights.push('📚 Wide Range of Courses');
    if ((c.specializations || []).length >= 3) highlights.push('🎯 Diverse Specializations');
    if (c.hostel?.available) highlights.push('🏠 On-Campus Hostel');
    if ((c.facilities || []).length >= 3) highlights.push('🏫 Well-Equipped Campus');
    if (c.type === 'Government' || c.type === 'Government-Aided') highlights.push('🏛️ Government-Backed Trust');
    if (!highlights.length) highlights.push('✅ Solid All-Round Choice');
    return { ...c, highlights };
  });

  res.json({ success: true, colleges: withHighlights });
});

app.get('/api/courses', cached('courses', async (req, res) => {
  const courses = await Course.find().select('name level durationYears').limit(500).lean();
  res.set('Cache-Control', 'public, max-age=300');
  res.json({ success: true, courses });
}));

/* ============================================================================
 *  COLLEGE INTEREST LEAD CAPTURE
 *  Frontend shows a short form (name/phone/email) before revealing the
 *  compare table or a college's full detail card. This captures which
 *  college(s) each visitor is interested in for the admin to follow up.
 *  Not tied to auth - works for logged-out visitors too, but attaches the
 *  student record automatically if they happen to be logged in.
 * ==========================================================================*/
app.post('/api/college-interest', validate({
  name: { required: true, minLength: 2 },
  phone: { required: true, minLength: 8 },
}), async (req, res) => {
  const { name, phone, email, collegeIds, context } = req.body;
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  let studentId;
  if (token) {
    try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* not logged in / invalid - fine, still capture the lead */ }
  }
  // collegeIds can include demo/mock IDs (e.g. "mock-rvce") when the public
  // pages are showing fallback sample data because no real colleges have
  // been added yet - filter those out instead of letting a bad ObjectId
  // crash the whole lead submission. The lead itself is still captured.
  const validCollegeIds = (Array.isArray(collegeIds) ? collegeIds : [])
    .filter((id) => mongoose.isValidObjectId(id))
    .slice(0, 4);
  const entry = await CollegeInterest.create({
    name, phone, email: email || '', student: studentId,
    colleges: validCollegeIds,
    context: context === 'view' ? 'view' : 'compare',
  });
  res.status(201).json({ success: true, message: 'Thanks! Here are your results.', interestId: entry._id });
});

/* ============================================================================
 *  REFERRAL SYSTEM (Career Assessment / Subscription)
 * ==========================================================================*/
app.get('/api/referrals/my', authRequired, async (req, res) => {
  if (req.user.role !== 'student') return res.status(403).json({ success: false, message: 'Students only' });
  const student = await Student.findById(req.user.id);
  const referrals = await Referral.find({ referringStudent: req.user.id }).sort({ createdAt: -1 });
  res.json({ success: true, referralCode: student.referralCode, referrals });
});

app.post('/api/referrals/track', async (req, res) => {
  // Called when someone lands on a referral link, before they register
  const { referralCode, name, email, type } = req.body;
  const referrer = await Student.findOne({ referralCode });
  if (!referrer) return res.status(404).json({ success: false, message: 'Invalid referral code' });
  const ref = await Referral.create({
    referringStudent: referrer._id, referralCode, referredName: name, referredEmail: email,
    type: type || 'assessment', status: 'pending',
  });
  res.json({ success: true, referral: ref });
});

/* ============================================================================
 *  COLLEGE ADMISSION REFERRAL (separate, independent link, admin-managed)
 * ==========================================================================*/
app.get('/api/college-referral/:code', async (req, res) => {
  const ref = await CollegeReferral.findOne({ referralLinkCode: req.params.code }).populate('college');
  if (!ref) return res.status(404).json({ success: false, message: 'Referral link not found' });
  const payload = ref.toObject();
  if (ref.hideCollegeFromReferredView) delete payload.college; // hide college details from referred customer
  res.json({ success: true, referral: payload });
});

app.post('/api/college-referral/:code/enquire', async (req, res) => {
  const { name, phone, email, message } = req.body;
  const ref = await CollegeReferral.findOne({ referralLinkCode: req.params.code });
  if (!ref) return res.status(404).json({ success: false, message: 'Referral link not found' });
  ref.enquiries.push({ name, phone, email, message });
  await ref.save();
  res.status(201).json({ success: true, message: 'Enquiry submitted successfully' });
});

/* ============================================================================
 *  CONTACT US (public)
 * ==========================================================================*/
app.post('/api/contact', validate({
  name: { required: true, minLength: 2 },
  email: { required: true, type: 'email' },
  message: { required: true, minLength: 5 },
}), async (req, res) => {
  const { name, email, phone, message } = req.body;
  const msg = await ContactMessage.create({ name, email, phone, message });
  res.status(201).json({ success: true, message: 'Message received, we will get back to you soon.', data: msg });
});

/* ============================================================================
 *  PUBLIC HOME PAGE CONTENT (slider + pages)
 * ==========================================================================*/
app.get('/api/sliders', cached('sliders', async (req, res) => {
  const sliders = await Slider.find({ isActive: true }).sort({ order: 1 }).lean();
  res.set('Cache-Control', 'public, max-age=60');
  res.json({ success: true, sliders });
}));

app.get('/api/pages/:slug', cached('page', async (req, res) => {
  const page = await PageContent.findOne({ slug: req.params.slug }).lean();
  res.set('Cache-Control', 'public, max-age=120');
  res.json({ success: true, page: page || null });
}));

/* ============================================================================
 *  ADMIN AUTH
 * ==========================================================================*/
app.post('/api/admin/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    const admin = await Admin.findOne({ email: (email || '').toLowerCase() });
    if (!admin) return res.status(401).json({ success: false, message: 'Invalid credentials' });
    const match = await bcrypt.compare(password, admin.password);
    if (!match) return res.status(401).json({ success: false, message: 'Invalid credentials' });
    const token = generateToken({ id: admin._id, role: 'admin' });
    res.json({ success: true, token, admin: { id: admin._id, name: admin.name, email: admin.email, role: admin.role } });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

/* ============================================================================
 *  ADMIN DASHBOARD (aggregated stats)
 * ==========================================================================*/
app.get('/api/admin/dashboard', adminOnly, async (req, res) => {
  const [studentCount, activeSubs, totalAssessments, totalRevenueAgg, pendingReferrals, openEnquiries] = await Promise.all([
    Student.countDocuments(),
    Subscription.countDocuments({ status: 'active', endDate: { $gte: new Date() } }),
    AssessmentResult.countDocuments(),
    Payment.aggregate([{ $match: { status: 'success' } }, { $group: { _id: null, total: { $sum: '$amount' } } }]),
    Referral.countDocuments({ status: 'pending' }),
    CollegeReferral.countDocuments({ status: 'open' }),
    Student.countDocuments({ 'kyc.status': 'pending' })
  ]);
  res.json({
    success: true,
    stats: {
      studentCount, activeSubscriptions: activeSubs, totalAssessments,
      totalRevenue: totalRevenueAgg[0]?.total || 0, pendingReferrals, openCollegeReferrals: openEnquiries,
    },
  });
});

/* ---- Manage Students ---- */
app.get('/api/admin/students', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.search) {
    const term = String(req.query.search).slice(0, 60);
    filter.$or = [
      { fullName: { $regex: term, $options: 'i' } },
      { email: { $regex: term, $options: 'i' } },
      { phone: { $regex: term, $options: 'i' } },
    ];
  }
  const [students, total] = await Promise.all([
    Student.find(filter)
      .select('-password -resetToken -resetTokenExpiry')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    Student.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: students, total, page, limit, extraKey: 'students' });
});
app.put('/api/admin/students/:id', adminOnly, async (req, res) => {
  const student = await Student.findByIdAndUpdate(req.params.id, req.body, { new: true }).select('-password');
  res.json({ success: true, student });
});
app.delete('/api/admin/students/:id', adminOnly, async (req, res) => {
  await Student.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'Student removed' });
});

/* ---- Manage Subscription Plans + purchased subscriptions ---- */
app.get('/api/admin/subscription-plans', adminOnly, async (req, res) => {
  const plans = await SubscriptionPlan.find();
  res.json({ success: true, plans });
});
app.post('/api/admin/subscription-plans', adminOnly, async (req, res) => {
  const plan = await SubscriptionPlan.create(req.body);
  cacheInvalidate('subscription-plans');
  res.status(201).json({ success: true, plan });
});
app.put('/api/admin/subscription-plans/:id', adminOnly, async (req, res) => {
  const plan = await SubscriptionPlan.findByIdAndUpdate(req.params.id, req.body, { new: true });
  cacheInvalidate('subscription-plans');
  res.json({ success: true, plan });
});
app.delete('/api/admin/subscription-plans/:id', adminOnly, async (req, res) => {
  await SubscriptionPlan.findByIdAndDelete(req.params.id);
  cacheInvalidate('subscription-plans');
  res.json({ success: true, message: 'Plan deleted' });
});
app.get('/api/admin/subscriptions', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  const [subs, total] = await Promise.all([
    Subscription.find(filter)
      .populate('student', 'fullName email').populate('plan', 'name price durationInDays')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    Subscription.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: subs, total, page, limit, extraKey: 'subscriptions' });
});

/* ---- Manage Career Assessments (question bank + results) ---- */
app.get('/api/admin/assessment-questions', adminOnly, async (req, res) => {
  const questions = await AssessmentQuestion.find();
  res.json({ success: true, questions });
});
app.post('/api/admin/assessment-questions', adminOnly, async (req, res) => {
  const question = await AssessmentQuestion.create(req.body);
  res.status(201).json({ success: true, question });
});
app.put('/api/admin/assessment-questions/:id', adminOnly, async (req, res) => {
  const question = await AssessmentQuestion.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json({ success: true, question });
});
app.delete('/api/admin/assessment-questions/:id', adminOnly, async (req, res) => {
  await AssessmentQuestion.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'Question deleted' });
});

// Bulk question upload via Excel/CSV, for admins who'd rather prep a
// spreadsheet than add questions one at a time. Expected columns (header
// row, case-insensitive, order doesn't matter):
//   Category, Question, AdaptiveForField,
//   Option1, Option1Weights, Option2, Option2Weights, Option3, Option3Weights, Option4, Option4Weights
// A "Weights" cell looks like:  Engineering & Technology:3, Science & Research:1
function parseWeightsCell(cell) {
  const weights = {};
  String(cell || '').split(',').forEach((pair) => {
    const [field, val] = pair.split(':');
    if (field && field.trim() && val !== undefined) weights[field.trim()] = Number(val.trim()) || 0;
  });
  return weights;
}
app.post('/api/admin/assessment-questions/upload-excel', adminOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'Excel file is required (field name: file)' });
    const rows = parseExcelRows(req.file.path);
    const validCategories = ['interest', 'aptitude', 'personality', 'adaptive'];
    const docs = [];
    const errors = [];
    rows.forEach((row, i) => {
      const category = String(row.category || '').trim().toLowerCase();
      const question = String(row.question || '').trim();
      if (!validCategories.includes(category) || !question) {
        errors.push(`Row ${i + 2}: missing/invalid category or question`);
        return;
      }
      const options = [];
      for (let n = 1; n <= 6; n += 1) {
        const label = row[`option${n}`];
        if (label && String(label).trim()) {
          options.push({ label: String(label).trim(), weights: parseWeightsCell(row[`option${n}weights`]) });
        }
      }
      if (options.length < 2) { errors.push(`Row ${i + 2}: needs at least 2 options`); return; }
      docs.push({ category, question, options, adaptiveForField: String(row.adaptiveforfield || '').trim() });
    });
    if (docs.length) await AssessmentQuestion.insertMany(docs);
    res.status(201).json({ success: true, imported: docs.length, skipped: errors.length, errors: errors.slice(0, 20) });
  } catch (err) {
    res.status(400).json({ success: false, message: 'Could not process file: ' + err.message });
  }
});
app.get('/api/admin/assessment-results', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const [results, total] = await Promise.all([
    AssessmentResult.find()
      .select('-answers') // answers array can be large; omit from list view, fetch per-record if needed
      .populate('student', 'fullName email')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    AssessmentResult.countDocuments(),
  ]);
  paginatedResponse(res, { data: results, total, page, limit, extraKey: 'results' });
});

/* ---- Manage Career Database (career fields + career values) ----
 * This is the "career map" behind every assessment report: each CareerField
 * carries the careers, courses and roadmap stages shown once a student's
 * answers point to it. Admins edit these directly - no code changes needed
 * to update the roadmap a student sees. */
app.get('/api/admin/career-fields', adminOnly, async (req, res) => {
  const fields = await CareerField.find().sort({ name: 1 }).lean();
  res.json({ success: true, fields });
});
app.post('/api/admin/career-fields', adminOnly, async (req, res) => {
  const field = await CareerField.create(req.body);
  cacheInvalidate('career-fields');
  res.status(201).json({ success: true, field });
});
app.put('/api/admin/career-fields/:id', adminOnly, async (req, res) => {
  const field = await CareerField.findByIdAndUpdate(req.params.id, req.body, { new: true, runValidators: true });
  cacheInvalidate('career-fields');
  res.json({ success: true, field });
});
app.delete('/api/admin/career-fields/:id', adminOnly, async (req, res) => {
  await CareerField.findByIdAndDelete(req.params.id);
  cacheInvalidate('career-fields');
  res.json({ success: true, message: 'Career field deleted' });
});

app.get('/api/admin/career-values', adminOnly, async (req, res) => {
  const values = await CareerValue.find().sort({ label: 1 }).lean();
  res.json({ success: true, values });
});
app.post('/api/admin/career-values', adminOnly, async (req, res) => {
  const value = await CareerValue.create(req.body);
  cacheInvalidate('career-values');
  res.status(201).json({ success: true, value });
});
app.put('/api/admin/career-values/:id', adminOnly, async (req, res) => {
  const value = await CareerValue.findByIdAndUpdate(req.params.id, req.body, { new: true, runValidators: true });
  cacheInvalidate('career-values');
  res.json({ success: true, value });
});
app.delete('/api/admin/career-values/:id', adminOnly, async (req, res) => {
  await CareerValue.findByIdAndDelete(req.params.id);
  cacheInvalidate('career-values');
  res.json({ success: true, message: 'Career value deleted' });
});

/* ---- Manage Payments & Transactions ---- */
app.get('/api/admin/payments', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  if (req.query.purpose) filter.purpose = req.query.purpose;
  const [payments, total] = await Promise.all([
    Payment.find(filter).populate('student', 'fullName email')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    Payment.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: payments, total, page, limit, extraKey: 'payments' });
});

// Bulk cleanup - remove all abandoned/failed payment attempts
// NOTE: declared BEFORE '/:id' so a literal path is always matched first.
app.delete('/api/admin/payments/cleanup/stale', adminOnly, async (req, res) => {
  try {
    const result = await Payment.deleteMany({ status: { $in: ['created', 'failed'] } });
    res.json({ success: true, deletedCount: result.deletedCount });
  } catch (err) {
    console.error('[DELETE /admin/payments/cleanup/stale]', err);
    res.status(500).json({ success: false, message: err.message || 'Failed to clear stale records' });
  }
});
 
app.delete('/api/admin/payments/:id', adminOnly, async (req, res) => {
  try {
    const payment = await Payment.findById(req.params.id);
    if (!payment) return res.status(404).json({ success: false, message: 'Payment not found' });
    if (payment.status === 'success') {
      return res.status(400).json({ success: false, message: 'Cannot delete a successful payment record' });
    }
    await Payment.findByIdAndDelete(req.params.id);
    res.json({ success: true, message: 'Payment record removed' });
  } catch (err) {
    console.error('[DELETE /admin/payments/:id]', err);
    res.status(500).json({ success: false, message: err.message || 'Failed to delete payment' });
  }
});



/* ---- Manage Assessment/Subscription Referrals ---- */
app.get('/api/admin/referrals', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.type) filter.type = req.query.type;
  if (req.query.status) filter.status = req.query.status;
  const [referrals, total] = await Promise.all([
    Referral.find(filter).populate('referringStudent', 'fullName email referralCode')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    Referral.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: referrals, total, page, limit, extraKey: 'referrals' });
});

/* ---- College Interest Leads (from compare/view lead-capture form) ---- */
app.get('/api/admin/college-interest', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  const [leads, total] = await Promise.all([
    CollegeInterest.find(filter)
      .populate('colleges', 'name location')
      .populate('student', 'fullName email')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    CollegeInterest.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: leads, total, page, limit, extraKey: 'leads' });
});
app.put('/api/admin/college-interest/:id', adminOnly, async (req, res) => {
  const lead = await CollegeInterest.findByIdAndUpdate(req.params.id, { status: req.body.status }, { new: true });
  res.json({ success: true, lead });
});

/* ============================================================================
 *  ADMIN: KYC VERIFICATION
 * ==========================================================================*/
function maskAccount(n) {
  const s = String(n || '');
  return s.length > 4 ? '•'.repeat(s.length - 4) + s.slice(-4) : s;
}
function docFlags(d = {}) {
  return { passbook: !!d.passbook, panCard: !!d.panCard, aadhaarCard: !!d.aadhaarCard };
}

// List: everyone who has submitted KYC. ?status=pending|verified|rejected|all, ?search=
app.get('/api/admin/kyc', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (['pending', 'verified', 'rejected'].includes(req.query.status)) filter['kyc.status'] = req.query.status;
  else filter['kyc.status'] = { $in: ['pending', 'verified', 'rejected'] };
  if (req.query.search) {
    const term = String(req.query.search).slice(0, 60);
    filter.$or = [
      { fullName: { $regex: term, $options: 'i' } },
      { email: { $regex: term, $options: 'i' } },
      { phone: { $regex: term, $options: 'i' } },
    ];
  }
  const [students, total, pending, verified, rejected] = await Promise.all([
    Student.find(filter)
      .select('fullName email phone kyc.status kyc.bankName kyc.accountNumber kyc.submittedAt kyc.documents')
      .sort({ 'kyc.submittedAt': -1 }).skip(skip).limit(limit).lean(),
    Student.countDocuments(filter),
    Student.countDocuments({ 'kyc.status': 'pending' }),
    Student.countDocuments({ 'kyc.status': 'verified' }),
    Student.countDocuments({ 'kyc.status': 'rejected' }),
  ]);
  res.json({
    success: true,
    students: students.map((s) => ({
      ...s,
      kyc: { ...s.kyc, accountNumber: maskAccount(s.kyc?.accountNumber), documents: docFlags(s.kyc?.documents) },
    })),
    counts: { pending, verified, rejected },
    pagination: { page, limit, total, totalPages: Math.max(1, Math.ceil(total / limit)) },
  });
});

// One submission, with full details (for the review window)
app.get('/api/admin/kyc/:id', adminOnly, async (req, res) => {
  if (!mongoose.isValidObjectId(req.params.id)) return res.status(400).json({ success: false, message: 'Invalid id' });
  const s = await Student.findById(req.params.id).select('fullName email phone kyc').lean();
  if (!s || !s.kyc || s.kyc.status === 'not_submitted') return res.status(404).json({ success: false, message: 'No KYC submission found' });
  res.json({ success: true, student: { ...s, kyc: { ...s.kyc, documents: docFlags(s.kyc.documents) } } });
});

// Stream one private document. Admin token required; never cached.
app.get('/api/admin/kyc/:id/document/:type', adminOnly, async (req, res) => {
  const { id, type } = req.params;
  if (!KYC_DOC_TYPES.includes(type) || !mongoose.isValidObjectId(id)) return res.status(400).json({ success: false, message: 'Invalid request' });
  const s = await Student.findById(id).select('kyc.documents').lean();
  const fileName = s?.kyc?.documents?.[type];
  if (!fileName) return res.status(404).json({ success: false, message: 'Document not uploaded' });
  const abs = path.join(kycDir, path.basename(fileName));
  if (!fs.existsSync(abs)) return res.status(404).json({ success: false, message: 'File missing on server' });
  res.set('Cache-Control', 'private, no-store');
  res.set('X-Content-Type-Options', 'nosniff');
  res.sendFile(abs);
});

// Verify or reject
app.put('/api/admin/kyc/:id', adminOnly, async (req, res) => {
  if (!mongoose.isValidObjectId(req.params.id)) return res.status(400).json({ success: false, message: 'Invalid id' });
  const { action, reason } = req.body;
  if (!['verify', 'reject'].includes(action)) return res.status(400).json({ success: false, message: 'action must be verify or reject' });
  if (action === 'reject' && !String(reason || '').trim()) {
    return res.status(400).json({ success: false, message: 'Please give a reason so the student knows what to fix' });
  }
  const s = await Student.findById(req.params.id).select('kyc').lean();
  if (!s || !s.kyc || s.kyc.status === 'not_submitted') return res.status(404).json({ success: false, message: 'No KYC submission found' });
  if (action === 'verify') {
    const d = s.kyc.documents || {};
    if (!KYC_DOC_TYPES.every((t) => d[t])) return res.status(400).json({ success: false, message: 'All three documents are needed before verifying' });
  }
  const student = await Student.findByIdAndUpdate(req.params.id, {
    $set: {
      'kyc.status': action === 'verify' ? 'verified' : 'rejected',
      'kyc.verifiedAt': action === 'verify' ? new Date() : null,
      'kyc.rejectionReason': action === 'verify' ? '' : String(reason).trim().slice(0, 300),
      'kyc.reviewedBy': req.user.id,
    },
  }, { new: true }).select('fullName kyc.status');
  res.json({ success: true, student });
});




/* ---- Manage College Admission Referrals ---- */
app.get('/api/admin/college-referrals', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  const [refs, total] = await Promise.all([
    CollegeReferral.find(filter).populate('college', 'name location')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    CollegeReferral.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: refs, total, page, limit, extraKey: 'referrals' });
});
app.post('/api/admin/college-referrals', adminOnly, async (req, res) => {
  let code;
  do { code = 'CR' + Math.floor(100000 + Math.random() * 900000); } while (await CollegeReferral.findOne({ referralLinkCode: code }));
  const ref = await CollegeReferral.create({ ...req.body, referralLinkCode: code });
  res.status(201).json({ success: true, referral: ref });
});
app.put('/api/admin/college-referrals/:id', adminOnly, async (req, res) => {
  const ref = await CollegeReferral.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json({ success: true, referral: ref });
});
app.delete('/api/admin/college-referrals/:id', adminOnly, async (req, res) => {
  await CollegeReferral.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'College referral deleted' });
});

/* ============================================================================
 *  UPLOAD & MANAGE KCET / PGCET CUTOFF DATA
 *  Supports: paginated listing, manual add, and bulk Excel (.xlsx/.xls) upload.
 *  Excel columns expected (header row, case-insensitive):
 *    KCET:  Year | College | Course | Category | Round | CutoffRank | 371J (Y/N)
 *    PGCET: Year | College | Course | Category | CollegeType | CutoffRank
 *  College/Course are matched by name (case-insensitive) and auto-created if
 *  they don't exist yet, so admins don't have to pre-create every row's college.
 * ==========================================================================*/
app.get('/api/admin/kcet-cutoffs', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.year) filter.year = Number(req.query.year);
  if (req.query.category) filter.category = req.query.category;
  const [data, total] = await Promise.all([
    KcetCutoff.find(filter)
      .populate('college', 'name location').populate('course', 'name level')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    KcetCutoff.countDocuments(filter),
  ]);
  paginatedResponse(res, { data, total, page, limit });
});
app.post('/api/admin/kcet-cutoffs', adminOnly, validate({
  year: { required: true, type: 'number' },
  category: { required: true },
  cutoffRank: { required: true, type: 'number', min: 1 },
}), async (req, res) => {
  // Accept college/course either as an ObjectId (from a dropdown) or a plain
  // name string (from the quick-add form) - resolve/auto-create by name either way.
  const body = { ...req.body };
  if ((body.college && !mongoose.isValidObjectId(body.college)) || (body.course && !mongoose.isValidObjectId(body.course))) {
    const { collegeMap, courseMap } = await resolveCollegeAndCourseIds([{ college: body.college, course: body.course }]);
    if (body.college && !mongoose.isValidObjectId(body.college)) body.college = collegeMap.get(String(body.college).toLowerCase());
    if (body.course && !mongoose.isValidObjectId(body.course)) body.course = courseMap.get(String(body.course).toLowerCase());
  }
  const record = await KcetCutoff.create(body);
  const populated = await KcetCutoff.findById(record._id).populate('college', 'name').populate('course', 'name').lean();
  res.status(201).json({ success: true, data: populated });
});
app.delete('/api/admin/kcet-cutoffs/:id', adminOnly, async (req, res) => {
  await KcetCutoff.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'Record deleted' });
});
app.delete('/api/admin/kcet-cutoffs', adminOnly, async (req, res) => {
  // Bulk clear (e.g. before re-uploading a fresh year's Excel sheet)
  const filter = {};
  if (req.query.year) filter.year = Number(req.query.year);
  const result = await KcetCutoff.deleteMany(filter);
  res.json({ success: true, deletedCount: result.deletedCount });
});
app.post('/api/admin/kcet-cutoffs/upload-excel', adminOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'Excel file is required (field name: file)' });
    const rows = parseExcelRows(req.file.path);
    const result = await bulkUpsertCutoffRows(rows, 'kcet');
    res.status(201).json({ success: true, ...result });
  } catch (err) {
    res.status(400).json({ success: false, message: 'Could not process Excel file: ' + err.message });
  }
});

// Reads a real KEA cutoff PDF (like the official "1st Round KCET Allotment
// Cut-Off Ranks" publication) straight into the system - no manual
// spreadsheet re-typing needed. College/Course records are matched by name
// and auto-created, exactly like the Excel path above.
app.post('/api/admin/kcet-cutoffs/upload-pdf', adminOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'PDF file is required (field name: file)' });
    const text = extractPdfTextViaPoppler(req.file.path);
    const yearMatch = text.match(/UGCET-(\d{4})|CET-(\d{4})/i);
    const year = Number(req.body.year) || Number(yearMatch?.[1] || yearMatch?.[2]) || new Date().getFullYear();
    const round = req.body.round || 'Round 1';
    const parsedRows = parseKeaCutoffPdfText(text).map((r) => ({ ...r, year, round }));
    if (!parsedRows.length) {
      return res.status(400).json({ success: false, message: 'No cutoff rows could be read from this PDF. Make sure it is the KEA "College: ... / Course Name ..." table format.' });
    }
    const result = await bulkUpsertCutoffRows(parsedRows, 'kcet');
    res.status(201).json({ success: true, year, ...result });
  } catch (err) {
    res.status(400).json({ success: false, message: err.message });
  }
});

app.get('/api/admin/pgcet-cutoffs', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.year) filter.year = Number(req.query.year);
  if (req.query.category) filter.category = req.query.category;
  const [data, total] = await Promise.all([
    PgcetCutoff.find(filter)
      .populate('college', 'name location').populate('course', 'name level')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    PgcetCutoff.countDocuments(filter),
  ]);
  paginatedResponse(res, { data, total, page, limit });
});
app.post('/api/admin/pgcet-cutoffs', adminOnly, validate({
  year: { required: true, type: 'number' },
  category: { required: true },
  cutoffRank: { required: true, type: 'number', min: 1 },
}), async (req, res) => {
  const body = { ...req.body };
  if ((body.college && !mongoose.isValidObjectId(body.college)) || (body.course && !mongoose.isValidObjectId(body.course))) {
    const { collegeMap, courseMap } = await resolveCollegeAndCourseIds([{ college: body.college, course: body.course }]);
    if (body.college && !mongoose.isValidObjectId(body.college)) body.college = collegeMap.get(String(body.college).toLowerCase());
    if (body.course && !mongoose.isValidObjectId(body.course)) body.course = courseMap.get(String(body.course).toLowerCase());
  }
  const record = await PgcetCutoff.create(body);
  const populated = await PgcetCutoff.findById(record._id).populate('college', 'name').populate('course', 'name').lean();
  res.status(201).json({ success: true, data: populated });
});
app.delete('/api/admin/pgcet-cutoffs/:id', adminOnly, async (req, res) => {
  await PgcetCutoff.findByIdAndDelete(req.params.id);
  res.json({ success: true, message: 'Record deleted' });
});
app.delete('/api/admin/pgcet-cutoffs', adminOnly, async (req, res) => {
  const filter = {};
  if (req.query.year) filter.year = Number(req.query.year);
  const result = await PgcetCutoff.deleteMany(filter);
  res.json({ success: true, deletedCount: result.deletedCount });
});
app.post('/api/admin/pgcet-cutoffs/upload-excel', adminOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'Excel file is required (field name: file)' });
    const rows = parseExcelRows(req.file.path);
    const result = await bulkUpsertCutoffRows(rows, 'pgcet');
    res.status(201).json({ success: true, ...result });
  } catch (err) {
    res.status(400).json({ success: false, message: 'Could not process Excel file: ' + err.message });
  }
});

// Same KEA-format PDF reader as the KCET route above, for PGCET cutoff publications.
app.post('/api/admin/pgcet-cutoffs/upload-pdf', adminOnly, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'PDF file is required (field name: file)' });
    const text = extractPdfTextViaPoppler(req.file.path);
    const yearMatch = text.match(/PGCET-(\d{4})|CET-(\d{4})/i);
    const year = Number(req.body.year) || Number(yearMatch?.[1] || yearMatch?.[2]) || new Date().getFullYear();
    const parsedRows = parseKeaCutoffPdfText(text).map((r) => ({ ...r, year, collegetype: req.body.collegeType || 'Private' }));
    if (!parsedRows.length) {
      return res.status(400).json({ success: false, message: 'No cutoff rows could be read from this PDF. Make sure it is the KEA "College: ... / Course Name ..." table format.' });
    }
    const result = await bulkUpsertCutoffRows(parsedRows, 'pgcet');
    res.status(201).json({ success: true, year, ...result });
  } catch (err) {
    res.status(400).json({ success: false, message: err.message });
  }
});

app.get('/api/admin/predictor-leads', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.examType) filter.examType = req.query.examType; // 'kcet' | 'pgcet'
  if (req.query.accessType) filter.accessType = req.query.accessType;
  if (req.query.search) {
    const term = String(req.query.search).slice(0, 60);
    filter.$or = [
      { name: { $regex: term, $options: 'i' } },
      { phone: { $regex: term, $options: 'i' } },
      { email: { $regex: term, $options: 'i' } },
    ];
  }
  const [leads, total] = await Promise.all([
    PredictorLead.find(filter).populate('paymentId', 'amount status').sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    PredictorLead.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: leads, total, page, limit, extraKey: 'leads' });
});






/* ---- Manage Colleges ----
 * One admin form covers every field the hub/compare page and detail page
 * show: basic info, accreditation, specializations, fees, placements,
 * hostel (2/3/4 share), facilities, contact and media. The frontend sends
 * comma-separated lists for array fields and JSON strings for nested
 * objects (fees/placements/hostel/contact) inside a multipart form so the
 * logo/cover/gallery files can ride along in the same request. */
const collegeUpload = upload.fields([
  { name: 'image', maxCount: 1 },
  { name: 'logo', maxCount: 1 },
  { name: 'gallery', maxCount: 8 },
]);

function toList(value) {
  if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
  if (typeof value === 'string') return value.split(',').map((v) => v.trim()).filter(Boolean);
  return [];
}

function parseJSONField(value, fallback = {}) {
  if (!value) return fallback;
  if (typeof value === 'object') return value;
  try { return JSON.parse(value); } catch { return fallback; }
}

function buildCollegePayload(body, files) {
  const payload = {
    name: body.name,
    code: body.code || '',
    location: body.location || '',
    state: body.state || '',
    type: body.type || 'Private',
    establishedYear: body.establishedYear ? Number(body.establishedYear) : undefined,
    affiliatedUniversity: body.affiliatedUniversity || '',
    about: body.about || '',
    description: body.description || '',
    rating: body.rating ? Number(body.rating) : 0,
    ranking: body.ranking ? Number(body.ranking) : undefined,
    accreditations: toList(body.accreditations),
       specializations: toList(body.specializations),
    programs: body.programs !== undefined ? toList(body.programs) : undefined,
        area: body.area !== undefined ? body.area : undefined,
    courseTypes: body.courseTypes !== undefined ? toList(body.courseTypes) : undefined,
    facilities: toList(body.facilities),
    brochureUrl: body.brochureUrl || '',
    featured: body.featured === 'true' || body.featured === true,
    fees: parseJSONField(body.fees, {}),
    placements: parseJSONField(body.placements, {}),
    hostel: parseJSONField(body.hostel, {}),
    contact: parseJSONField(body.contact, {}),
  };
  if (body.coursesOffered) {
    payload.coursesOffered = typeof body.coursesOffered === 'string'
      ? body.coursesOffered.split(',').map((v) => v.trim()).filter(Boolean)
      : body.coursesOffered;
  }
  if (files?.image?.[0]) payload.image = '/uploads/' + files.image[0].filename;
  if (files?.logo?.[0]) payload.logo = '/uploads/' + files.logo[0].filename;
  if (files?.gallery?.length) payload.gallery = files.gallery.map((f) => '/uploads/' + f.filename);
  Object.keys(payload).forEach((k) => payload[k] === undefined && delete payload[k]);
  return payload;
}

app.post('/api/admin/colleges', adminOnly, collegeUpload, async (req, res) => {
  const payload = buildCollegePayload(req.body, req.files);
  const college = await College.create(payload);
  cacheInvalidate('colleges');
  cacheInvalidate('college-filters');
  res.status(201).json({ success: true, college });
});
app.put('/api/admin/colleges/:id', adminOnly, collegeUpload, async (req, res) => {
  const payload = buildCollegePayload(req.body, req.files);
  // Gallery uploads are additive (append new photos) unless the admin explicitly clears them.
  if (req.files?.gallery?.length && req.body.replaceGallery !== 'true') {
    const existing = await College.findById(req.params.id).select('gallery').lean();
    payload.gallery = [...(existing?.gallery || []), ...payload.gallery];
  }
  const college = await College.findByIdAndUpdate(req.params.id, payload, { new: true });
  cacheInvalidate('colleges');
  cacheInvalidate('college-filters');
  res.json({ success: true, college });
});
app.delete('/api/admin/colleges/:id', adminOnly, async (req, res) => {
  await College.findByIdAndDelete(req.params.id);
  cacheInvalidate('colleges');
  cacheInvalidate('college-filters');
  res.json({ success: true, message: 'College deleted' });
});

/* ---- Manage Courses ---- */
app.post('/api/admin/courses', adminOnly, async (req, res) => {
  const course = await Course.create(req.body);
  cacheInvalidate('courses');
  res.status(201).json({ success: true, course });
});
app.put('/api/admin/courses/:id', adminOnly, async (req, res) => {
  const course = await Course.findByIdAndUpdate(req.params.id, req.body, { new: true });
  cacheInvalidate('courses');
  res.json({ success: true, course });
});
app.delete('/api/admin/courses/:id', adminOnly, async (req, res) => {
  await Course.findByIdAndDelete(req.params.id);
  cacheInvalidate('courses');
  res.json({ success: true, message: 'Course deleted' });
});
app.delete('/api/admin/courses/cleanup/dirty', adminOnly, async (req, res) => {
  const result = await Course.deleteMany({ name: /Generated on|College:|Page\s+of/i });
  cacheInvalidate('kcet-meta');
  cacheInvalidate('pgcet-meta');
  res.json({ success: true, deletedCount: result.deletedCount });
});

/* ---- Manage Slider / Banners ---- */
app.post('/api/admin/sliders', adminOnly, upload.single('image'), async (req, res) => {
  const payload = { ...req.body };
  if (req.file) payload.image = '/uploads/' + req.file.filename;
  const slider = await Slider.create(payload);
  cacheInvalidate('sliders');
  res.status(201).json({ success: true, slider });
});
app.put('/api/admin/sliders/:id', adminOnly, upload.single('image'), async (req, res) => {
  const payload = { ...req.body };
  if (req.file) payload.image = '/uploads/' + req.file.filename;
  const slider = await Slider.findByIdAndUpdate(req.params.id, payload, { new: true });
  cacheInvalidate('sliders');
  res.json({ success: true, slider });
});
app.delete('/api/admin/sliders/:id', adminOnly, async (req, res) => {
  await Slider.findByIdAndDelete(req.params.id);
  cacheInvalidate('sliders');
  res.json({ success: true, message: 'Slide deleted' });
});

/* ---- Manage Website Pages & Content ---- */
app.get('/api/admin/pages', adminOnly, async (req, res) => {
  const pages = await PageContent.find();
  res.json({ success: true, pages });
});
app.put('/api/admin/pages/:slug', adminOnly, async (req, res) => {
  const page = await PageContent.findOneAndUpdate(
    { slug: req.params.slug }, { $set: req.body }, { new: true, upsert: true }
  );
  cacheInvalidate('page:');
  res.json({ success: true, page });
});

/* ---- Reports & Analytics ---- */
app.get('/api/admin/reports/summary', adminOnly, async (req, res) => {
  const [studentsByMonth, revenueByMonth, assessmentsByMonth] = await Promise.all([
    Student.aggregate([{ $group: { _id: { $dateToString: { format: '%Y-%m', date: '$createdAt' } }, count: { $sum: 1 } } }, { $sort: { _id: 1 } }]),
    Payment.aggregate([{ $match: { status: 'success' } }, { $group: { _id: { $dateToString: { format: '%Y-%m', date: '$createdAt' } }, total: { $sum: '$amount' } } }, { $sort: { _id: 1 } }]),
    AssessmentResult.aggregate([{ $group: { _id: { $dateToString: { format: '%Y-%m', date: '$createdAt' } }, count: { $sum: 1 } } }, { $sort: { _id: 1 } }]),
  ]);
  res.json({ success: true, studentsByMonth, revenueByMonth, assessmentsByMonth });
});

/* ---- Export Data (simple JSON export; swap to CSV lib if needed) ---- */
app.get('/api/admin/export/:collection', adminOnly, async (req, res) => {
  const map = { students: Student, payments: Payment, subscriptions: Subscription, assessments: AssessmentResult, colleges: College, referrals: Referral };
  const Model = map[req.params.collection];
  if (!Model) return res.status(400).json({ success: false, message: 'Unknown collection' });
  // Capped export - for very large datasets, use the paginated admin list endpoints
  // and export in batches instead of one unbounded dump.
  const data = await Model.find().sort({ createdAt: -1 }).limit(MAX_EXPORT_ROWS).lean();
  res.setHeader('Content-Disposition', `attachment; filename=${req.params.collection}.json`);
  res.json({ exportedCount: data.length, cappedAt: MAX_EXPORT_ROWS, data });
});

/* ---- Admin Profile & Settings ---- */
app.get('/api/admin/me', adminOnly, async (req, res) => {
  const admin = await Admin.findById(req.user.id).select('-password');
  res.json({ success: true, admin });
});
app.put('/api/admin/me', adminOnly, async (req, res) => {
  const updates = { ...req.body };
  if (updates.password) updates.password = await bcrypt.hash(updates.password, 10);
  const admin = await Admin.findByIdAndUpdate(req.user.id, updates, { new: true }).select('-password');
  res.json({ success: true, admin });
});

/* ---- Contact messages (admin view) ---- */
app.get('/api/admin/contact-messages', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  const [messages, total] = await Promise.all([
    ContactMessage.find(filter).sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    ContactMessage.countDocuments(filter),
  ]);
  paginatedResponse(res, { data: messages, total, page, limit, extraKey: 'messages' });
});
app.put('/api/admin/contact-messages/:id', adminOnly, async (req, res) => {
  const msg = await ContactMessage.findByIdAndUpdate(req.params.id, req.body, { new: true });
  res.json({ success: true, message: msg });
});


app.get('/api/reviews', cached('reviews', async (req, res) => {
  const limit = Math.min(Number(req.query.limit) || 12, 50);
  const filter = { status: 'approved' };
  if (req.query.source) filter.source = req.query.source;
 
  const [reviews, count, avgAgg] = await Promise.all([
    Review.find(filter)
      .select('displayName course city source rating review featured createdAt')
      .sort({ featured: -1, order: 1, createdAt: -1 })
      .limit(limit)
      .lean(),
    Review.countDocuments({ status: 'approved' }),
    Review.aggregate([
      { $match: { status: 'approved' } },
      { $group: { _id: null, avg: { $avg: '$rating' } } },
    ]),
  ]);
 
  res.set('Cache-Control', 'public, max-age=120');
  res.json({
    success: true,
    reviews: reviews.map((r) => ({
      id: String(r._id),
      name: r.displayName,
      course: [r.course, r.city].filter(Boolean).join(' · '),
      source: r.source === 'Other' ? '' : r.source,
      rating: r.rating,
      review: r.review,
      featured: !!r.featured,
    })),
    totalCount: count,
    averageRating: avgAgg[0]?.avg ? Number(avgAgg[0].avg.toFixed(1)) : null,
  });
}));
 
/* ---- Public: submit a review (goes to the moderation queue) ---- */
app.post('/api/reviews', validate({
  fullName: { required: true, minLength: 2 },
  rating: { required: true, type: 'number', min: 1 },
  review: { required: true, minLength: 15 },
}), async (req, res) => {
  try {
    const { fullName, email, phone, course, city, source, rating, review, consentToPublish } = req.body;
 
    if (consentToPublish !== true && consentToPublish !== 'true') {
      return res.status(400).json({
        success: false,
        message: 'We need your permission before we can publish your review.',
      });
    }
    if (Number(rating) < 1 || Number(rating) > 5) {
      return res.status(400).json({ success: false, message: 'Rating must be between 1 and 5' });
    }
    if (String(review).trim().length > 1000) {
      return res.status(400).json({ success: false, message: 'Please keep your review under 1000 characters' });
    }
 
    // Attach the student record if they happen to be logged in
    let studentId;
    const header = req.headers.authorization || '';
    const token = header.startsWith('Bearer ') ? header.slice(7) : null;
    if (token) {
      try { studentId = jwt.verify(token, JWT_SECRET).id; } catch { /* guest review, fine */ }
    }
 
    // Light anti-spam: one review per email or phone per 24h
    if (email || phone) {
      const recent = await Review.findOne({
        $or: [email ? { email: String(email).toLowerCase() } : null, phone ? { phone } : null].filter(Boolean),
        createdAt: { $gte: new Date(Date.now() - 24 * 60 * 60 * 1000) },
      }).select('_id').lean();
      if (recent) {
        return res.status(429).json({ success: false, message: 'Looks like you already sent us a review today — thank you!' });
      }
    }
 
    const validSources = ['KCET', 'PGCET', 'Assessment', 'Counselling', 'Colleges', 'Other'];
 
    const created = await Review.create({
      fullName: String(fullName).trim(),
      displayName: toDisplayName(fullName),
      email: email || '',
      phone: phone || '',
      course: (course || '').trim(),
      city: (city || '').trim(),
      source: validSources.includes(source) ? source : 'Other',
      rating: Number(rating),
      review: String(review).trim(),
      student: studentId,
      consentToPublish: true,
      consentAt: new Date(),
      status: 'pending',
    });
 
    res.status(201).json({
      success: true,
      message: 'Thank you! Your review has been sent for a quick check before it goes live.',
      reviewId: created._id,
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});
 
/* ---- Admin: moderation queue ---- */
app.get('/api/admin/reviews', adminOnly, async (req, res) => {
  const { page, limit, skip } = getPagination(req);
  const filter = {};
  if (req.query.status) filter.status = req.query.status;
  if (req.query.source) filter.source = req.query.source;
  if (req.query.search) {
    const term = String(req.query.search).slice(0, 60);
    filter.$or = [
      { fullName: { $regex: term, $options: 'i' } },
      { email: { $regex: term, $options: 'i' } },
      { review: { $regex: term, $options: 'i' } },
    ];
  }
 
  const [reviews, total, pendingCount] = await Promise.all([
    Review.find(filter).populate('student', 'fullName email')
      .sort({ createdAt: -1 }).skip(skip).limit(limit).lean(),
    Review.countDocuments(filter),
    Review.countDocuments({ status: 'pending' }),
  ]);
 
  res.json({
    success: true,
    reviews,
    pendingCount,
    pagination: { page, limit, total, totalPages: Math.max(1, Math.ceil(total / limit)) },
  });
});
 
// Approve / reject / feature / reorder / edit the display name
app.put('/api/admin/reviews/:id', adminOnly, async (req, res) => {
  const allowed = ['status', 'featured', 'order', 'displayName', 'course', 'city', 'source', 'adminNote', 'review'];
  const updates = {};
  allowed.forEach((k) => { if (req.body[k] !== undefined) updates[k] = req.body[k]; });
  if (updates.status === 'approved') updates.approvedAt = new Date();
 
  // Only one featured review at a time — it's the big pull-quote slot.
  if (updates.featured === true || updates.featured === 'true') {
    await Review.updateMany({ _id: { $ne: req.params.id } }, { featured: false });
  }
 
  const review = await Review.findByIdAndUpdate(req.params.id, updates, { new: true });
  cacheInvalidate('reviews');
  res.json({ success: true, review });
});
 
app.delete('/api/admin/reviews/:id', adminOnly, async (req, res) => {
  await Review.findByIdAndDelete(req.params.id);
  cacheInvalidate('reviews');
  res.json({ success: true, message: 'Review deleted' });
});
 



/* ============================================================================
 *  404 + ERROR HANDLER
 * ==========================================================================*/
app.use('/api', (req, res) => res.status(404).json({ success: false, message: 'API route not found' }));

app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ success: false, message: 'Server error', error: NODE_ENV === 'development' ? err.message : undefined });
});

/* ============================================================================
 *  START SERVER
 * ==========================================================================*/
app.listen(PORT, () => {
  console.log(`==================================================`);
  console.log(` MapMyCareer360 API running in ${NODE_ENV} mode`);
  console.log(` http://localhost:${PORT}`);
  console.log(`==================================================`);
});