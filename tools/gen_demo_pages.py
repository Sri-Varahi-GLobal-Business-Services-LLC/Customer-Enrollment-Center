# -*- coding: utf-8 -*-
"""Generate three Enrollment Center demo pages: combined (index), work center only, move-in only."""
import io, os

TEMPLATE = r"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title>
<style>
  :root{
    --sap-blue:#0070f2; --sap-blue-dark:#0057d2; --sap-shell:#12233f;
    --bg:#f5f6f7; --card:#ffffff; --border:#d9d9d9; --txt:#1d2d3e; --muted:#556b82;
    --green:#188918; --red:#d20a0a; --orange:#c35500; --chip-green:#ebf5eb; --chip-red:#ffeaea; --chip-orange:#fff3e6;
    --radius:10px; --shadow:0 1px 4px rgba(29,45,62,.12);
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--txt);font-size:14px;}
  .shell{background:var(--sap-shell);color:#fff;padding:10px 20px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
  .shell .logo{font-weight:700;font-size:16px}
  .shell .logo span{color:#7bb8ff}
  .shell .crumb{opacity:.75;font-size:12px}
  .shell .right{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:12px}
  .badge-demo{background:#ffb300;color:#3b2a00;font-weight:700;padding:2px 8px;border-radius:10px;font-size:11px}
  .badge-mode{background:var(--sap-blue);color:#fff;font-weight:700;padding:3px 12px;border-radius:14px;font-size:12px}
  .mode-toggle{display:flex;border:1px solid #5b7aa5;border-radius:16px;overflow:hidden}
  .mode-toggle button{background:transparent;color:#cfe0f5;border:0;padding:4px 12px;cursor:pointer;font-size:12px}
  .mode-toggle button.active{background:var(--sap-blue);color:#fff;font-weight:600}
  .custsel{background:#1d3358;color:#fff;border:1px solid #5b7aa5;border-radius:8px;padding:4px 8px;font-size:12px;max-width:280px}
  .ctx{background:var(--card);border-bottom:1px solid var(--border);padding:12px 20px;display:flex;gap:28px;flex-wrap:wrap;align-items:center}
  .ctx .avatar{width:42px;height:42px;border-radius:50%;background:var(--sap-blue);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:15px}
  .ctx .kv{font-size:12px;color:var(--muted)}
  .ctx .kv b{display:block;color:var(--txt);font-size:13px;font-weight:600}
  .ctx .name{font-size:16px;font-weight:700}
  .pill{display:inline-block;padding:1px 8px;border-radius:9px;font-size:11px;font-weight:600}
  .pill.blue{background:#e3f0ff;color:var(--sap-blue-dark)}
  .pill.orange{background:var(--chip-orange);color:var(--orange)}
  .pill.grey{background:#eceff2;color:var(--muted)}
  .wrap{max-width:1380px;margin:14px auto;padding:0 16px;display:grid;grid-template-columns:1fr 1fr;gap:14px}
  @media(max-width:980px){.wrap{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
  .card h2{font-size:14px;font-weight:700;padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px}
  .card h2 .cnt{background:#e8edf2;border-radius:9px;padding:0 8px;font-size:11px;color:var(--muted)}
  .card .body{padding:8px 16px 14px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{text-align:left;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px;padding:8px 8px 4px 0;border-bottom:1px solid var(--border)}
  td{padding:9px 8px 9px 0;border-bottom:1px solid #eef1f4;vertical-align:top}
  tr:last-child td{border-bottom:0}
  a.prog{color:var(--sap-blue-dark);font-weight:600;text-decoration:none;cursor:pointer}
  a.prog:hover{text-decoration:underline}
  .status{display:inline-block;padding:2px 9px;border-radius:9px;font-size:11px;font-weight:700}
  .st-enrolled{background:var(--chip-green);color:var(--green)}
  .st-pending{background:#e3f0ff;color:var(--sap-blue-dark)}
  .st-request{background:#f0eafd;color:#6c32a9}
  .st-rejected,.st-error{background:var(--chip-red);color:var(--red)}
  .st-approval{background:var(--chip-orange);color:var(--orange)}
  .st-deenrolled{background:#eceff2;color:var(--muted)}
  .sev{display:inline-block;padding:1px 8px;border-radius:4px;font-size:10.5px;font-weight:800;letter-spacing:.4px}
  .sev.hard{background:var(--red);color:#fff}
  .sev.warn{background:#ffb300;color:#3b2a00}
  .sev.info{background:#e3f0ff;color:var(--sap-blue-dark)}
  .btn{border:1px solid var(--sap-blue);background:#fff;color:var(--sap-blue-dark);border-radius:8px;padding:5px 12px;font-size:12px;font-weight:600;cursor:pointer}
  .btn:hover{background:#eaf3ff}
  .btn.primary{background:var(--sap-blue);color:#fff}
  .btn.primary:hover{background:var(--sap-blue-dark)}
  .btn.ghost{border-color:var(--border);color:var(--muted)}
  .btn:disabled{opacity:.45;cursor:not-allowed}
  .muted{color:var(--muted);font-size:12px}
  .empty{padding:22px 0;text-align:center;color:var(--muted);font-size:13px}
  .logwrap{max-width:1380px;margin:0 auto 20px;padding:0 16px}
  .log{background:#0f1c30;color:#c9d8ee;border-radius:var(--radius);font-family:Consolas,monospace;font-size:12px;padding:12px 16px;max-height:190px;overflow:auto}
  .log .t{color:#6f87a8}
  .log .api{color:#7bd88f}
  .log .evt{color:#ffb300}
  .log .err{color:#ff7a7a}
  .loghead{display:flex;align-items:center;gap:10px;padding:10px 4px 6px}
  .loghead h3{font-size:13px}
  .overlay{position:fixed;inset:0;background:rgba(18,35,63,.45);display:none;align-items:center;justify-content:center;z-index:50;padding:16px}
  .overlay.open{display:flex}
  .modal{background:#fff;border-radius:12px;max-width:560px;width:100%;max-height:88vh;overflow:auto;box-shadow:0 12px 40px rgba(0,0,0,.3)}
  .modal header{padding:14px 20px;border-bottom:1px solid var(--border);font-weight:700;font-size:15px;display:flex;justify-content:space-between;align-items:center}
  .modal .x{cursor:pointer;border:0;background:none;font-size:18px;color:var(--muted)}
  .modal .mbody{padding:16px 20px}
  .modal footer{padding:12px 20px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:10px}
  .field{margin-bottom:12px}
  .field label{display:block;font-size:12px;color:var(--muted);margin-bottom:4px;font-weight:600}
  .field input,.field select,.field textarea{width:100%;padding:7px 10px;border:1px solid var(--border);border-radius:8px;font-size:13px;font-family:inherit}
  .field .req{color:var(--red)}
  .consent{display:flex;gap:8px;align-items:flex-start;background:#f4f8fd;border:1px solid #d5e6fb;border-radius:8px;padding:10px;font-size:12.5px}
  .notice{border-radius:8px;padding:10px 12px;font-size:12.5px;margin-bottom:12px}
  .notice.warn{background:var(--chip-orange);border:1px solid #f0c08a;color:#7a4400}
  .notice.err{background:var(--chip-red);border:1px solid #f2b5b5;color:#8f0000}
  .notice.ok{background:var(--chip-green);border:1px solid #bfe3bf;color:#0f5c0f}
  .notice.info{background:#e9f2fd;border:1px solid #c2dcfa;color:#0b4ea2}
  .steps{display:flex;gap:0;margin:2px 0 14px}
  .steps div{flex:1;text-align:center;font-size:11px;color:var(--muted);padding:6px 2px;border-bottom:3px solid #e2e7ec}
  .steps div.on{color:var(--sap-blue-dark);font-weight:700;border-color:var(--sap-blue)}
  .toast{position:fixed;bottom:22px;right:22px;background:#12233f;color:#fff;border-radius:10px;padding:12px 18px;font-size:13px;box-shadow:0 8px 24px rgba(0,0,0,.35);opacity:0;transform:translateY(8px);transition:.25s;z-index:99;max-width:380px}
  .toast.show{opacity:1;transform:none}
  .movein-banner{max-width:1380px;margin:12px auto 0;padding:0 16px;display:none}
  .movein-banner.on{display:block}
  .movein-inner{background:#fff8e6;border:1px solid #f0d48a;border-radius:var(--radius);padding:10px 16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;font-size:13px}
  footer.page{max-width:1380px;margin:0 auto 30px;padding:0 16px;color:var(--muted);font-size:11.5px;line-height:1.6}
</style>
</head>
<body>

<div class="shell">
  <div class="logo">SAP <span>Service Cloud V2</span></div>
  <div class="crumb" id="crumb"></div>
  <div class="right">
    <span class="badge-demo">DEMO / SANDBOX</span>
    <span id="modeControls"></span>
    <select class="custsel" id="custSel" onchange="setCustomer(this.value)" title="Simulate identified customer"></select>
    <span>Agent: <b>A.Rivera</b></span>
  </div>
</div>

<div class="ctx" id="ctxbar"></div>

<div class="movein-banner" id="moveinBanner">
  <div class="movein-inner">
    <b>&#9888; Pre-active context:</b> Move-In document <b>MI-2026-00811</b> &mdash; contract not yet active.
    Enrollments captured here are stored as <span class="status st-request">Enrollment Request</span> and fulfilled automatically on contract activation.
    <button class="btn primary" style="margin-left:auto" onclick="simulateMoveInComplete()">&#9654; Simulate: Move-In Completed event</button>
    <button class="btn ghost" onclick="simulateMoveInCancel()">&#10005; Simulate: Move-In Cancelled</button>
  </div>
</div>

<div class="wrap">
  <div class="card">
    <h2>&#128203; Participating Programs <span class="cnt" id="cnt-part"></span></h2>
    <div class="body"><div id="panel-part"></div></div>
  </div>
  <div class="card">
    <h2>&#9989; Available (Eligible) Programs <span class="cnt" id="cnt-avail"></span></h2>
    <div class="body"><div id="panel-avail"></div></div>
  </div>
  <div class="card">
    <h2>&#9888;&#65039; Eligibility Messages <span class="cnt" id="cnt-msg"></span></h2>
    <div class="body"><div id="panel-msg"></div></div>
  </div>
  <div class="card">
    <h2>&#128344; Program History <span class="cnt" id="cnt-hist"></span></h2>
    <div class="body"><div id="panel-hist"></div></div>
  </div>
</div>

<div class="logwrap">
  <div class="loghead">
    <h3>&#128268; Integration Monitor &mdash; API calls &amp; Event Mesh (simulated)</h3>
    <button class="btn ghost" style="margin-left:auto" onclick="LOG.innerHTML=''">Clear</button>
  </div>
  <div class="log" id="log"></div>
</div>

<footer class="page">
  <b>Demo environment.</b> This static page simulates the BTP-hosted Enrollment Center described in the
  <i>Customer Enrollment Center for SAP Service Cloud V2 &mdash; Solution Design v2.0</i>.
  In the target architecture this UI is a SAPUI5 app on SAP BTP embedded in SC V2 (work center mashup / Move-In guided step).
  Program catalog drawn from the legacy Customer Programs Enrollment Center scope (PTR, CPP, CSDD, DPP, PrePay, Peak Power Savers, CARE/FERA)
  and the U.S. Utility Customer Programs Reference Catalog (DER/DR, energy efficiency, billing, payment, assistance and sustainability programs).
  All customers and data are fictitious; no real systems are called; data resets on page reload.
</footer>

<div class="overlay" id="overlay"><div class="modal" id="modal"></div></div>
<div class="toast" id="toast"></div>

<script>
/* =================== PAGE MODE =================== */
const PAGE_MODE = "__PAGEMODE__";   /* combined | standalone | movein */

/* =================== PROGRAM CATALOG (from legacy EC scope) =================== */
const CATALOG = {
  PTR:  {name:"Peak-Time Rebate (PTR)", cat:"Demand Response / Rate", recipe:"RECIPE_RATE",
         terms:"Earn bill credits by reducing usage during called peak events (max 12 events/season). Requires an AMI meter and residential rate class. Credits appear on the next bill cycle.",
         fields:[{k:"startDate",l:"Requested Start Date",t:"date"},{k:"notifChannel",l:"Event Notification Channel",t:"select",opts:["SMS","Email","Voice call"]}]},
  CPP:  {name:"Critical Peak Pricing (CPP)", cat:"Demand Response / Rate", recipe:"RECIPE_RATE",
         terms:"Discounted base rates in exchange for significantly higher prices during called critical peak events (max 15 events/year). Requires an AMI meter. Rate change effective next billing cycle.",
         fields:[{k:"startDate",l:"Requested Start Date",t:"date"},{k:"notifChannel",l:"Event Notification Channel",t:"select",opts:["SMS","Email","Voice call"]}]},
  TOU:  {name:"Time-of-Use Rate (TOU)", cat:"Demand Response / Rate", recipe:"RECIPE_RATE",
         terms:"Lower rates overnight and weekends, higher 3-8 PM weekdays. 12-month minimum participation. Rate change effective next billing cycle.",
         fields:[{k:"startDate",l:"Requested Start Date",t:"date"},{k:"ratePlan",l:"TOU Rate Plan",t:"select",opts:["TOU-Standard","TOU-EV Owner","TOU-Solar","TOU-Commercial"]}]},
  PPS:  {name:"Peak Power Savers (AC Cycling)", cat:"Demand Response / Device", recipe:"RECIPE_DEVICE",
         terms:"Utility-installed switch cycles the air-conditioning compressor during peak events (choose 50% or 100% cycling). Seasonal bill credit; override allowed twice per season.",
         fields:[{k:"cyclingOption",l:"Cycling Option",t:"select",opts:["50% cycling","100% cycling"]},{k:"acUnits",l:"Number of AC Units",t:"number",min:1,max:5}]},
  THERM:{name:"Smart Thermostat Program", cat:"Device / Service", recipe:"RECIPE_DEVICE",
         terms:"Free smart thermostat plus installation. Utility may adjust setpoint by up to 2 degrees during peak events (override allowed twice per season).",
         fields:[{k:"deviceModel",l:"Thermostat Model",t:"select",opts:["EcoStat 3","EcoStat 3 Pro","ComfortHub X"]},{k:"installDate",l:"Preferred Installation Date",t:"date"},{k:"homeWifi",l:"Wi-Fi available at premise?",t:"select",opts:["Yes","No"]}]},
  CSDD: {name:"Customer Selects Due Date (CSDD)", cat:"Billing Preference", recipe:"RECIPE_CSDD",
         terms:"Pick a fixed monthly bill due date (1st-28th). One change per 12 months.",
         fields:[{k:"dueDay",l:"Preferred Due Day of Month (1-28)",t:"number",min:1,max:28}]},
  EBILL:{name:"eBill / Paperless Billing", cat:"Billing Preference", recipe:"RECIPE_EBILL",
         terms:"Bills delivered electronically with email notification; paper bill suppressed. Can be combined with any payment program.",
         fields:[{k:"email",l:"Delivery Email Address",t:"text"}]},
  DPA:  {name:"Payment Plan / Deferred Payment Arrangement (DPA)", cat:"Payment", recipe:"RECIPE_DPP",
         terms:"Spread an outstanding balance over equal monthly installments billed alongside the current charges. Dunning is suppressed while the plan is kept; default terminates the plan and reinstates collections.",
         fields:[{k:"downPayment",l:"Down Payment",t:"select",opts:["0%","25%","50%"]},{k:"numInstallments",l:"Number of Installments",t:"select",opts:["6","12","18"]},{k:"startDate",l:"First Installment Date",t:"date"}]},
  DPP:  {name:"Dynamic Peak Pricing (DPP)", cat:"Demand Response / Rate", recipe:"RECIPE_RATE",
         terms:"Time-varying rate: lower base prices year-round in exchange for significantly higher prices during called peak events (typically summer weekday afternoons). Requires an AMI meter. Fulfilled as a rate/product change on the utilities contract (legacy: product change to the residential DPP/RCPP rate).",
         fields:[{k:"startDate",l:"Requested Start Date",t:"date"},{k:"notifChannel",l:"Event Notification Channel",t:"select",opts:["SMS","Email","Voice call"]}]},
  BB:   {name:"Budget Billing", cat:"Payment", recipe:"RECIPE_BB",
         terms:"Levelized monthly payment based on the 12-month rolling average; trued up at annual review. Requires 12 months of billing history at the premise.",
         fields:[{k:"startDate",l:"Requested Start Date",t:"date"}]},
  PREPAY:{name:"PrePay (Prepayment Program)", cat:"Payment", recipe:"RECIPE_PREPAY",
         terms:"Pay-as-you-go service: no deposit, no late fees. Balance decremented daily from usage; low-balance alerts at your chosen threshold; disconnection at zero balance per tariff. Not available for premises with an active Medical Alert / life-support flag.",
         fields:[{k:"initialCredit",l:"Initial Credit Purchase",t:"select",opts:["$25","$50","$100"]},{k:"lowBalanceThreshold",l:"Low-Balance Alert Threshold",t:"select",opts:["$10","$15","$25"]},{k:"notifChannel",l:"Alert Channel",t:"select",opts:["SMS","Email","Voice call"]}]},
  CARE: {name:"CARE / FERA Income-Qualified Discount", cat:"Assistance", recipe:"RECIPE_ASSIST",
         terms:"Monthly bill discount for income-qualified households (CARE ~30%, FERA ~18%). Self-certification at enrollment; post-enrollment income verification may be requested. Residential rate class only.",
         fields:[{k:"program",l:"Program",t:"select",opts:["CARE","FERA"]},{k:"householdSize",l:"Household Size",t:"number",min:1,max:12},{k:"incomeBracket",l:"Annual Household Income",t:"select",opts:["Below $40,000","$40,000-$55,000","$55,000-$70,000"]}]},
  MED:  {name:"Medical Alert Program", cat:"Assistance", recipe:"RECIPE_ASSIST",
         terms:"Flags the account for life-support / medical equipment at premise: advance outage notification and dunning protection. Physician certification required within 30 days.",
         fields:[{k:"contactPhone",l:"24h Contact Phone",t:"text"},{k:"equipment",l:"Medical Equipment Type",t:"select",opts:["Oxygen concentrator","Dialysis","CPAP/Ventilator","Other"]}]},
  CDR:  {name:"Commercial Demand Response (C&I)", cat:"Demand Response / Commercial", recipe:"RECIPE_RATE",
         terms:"Curtailment credits for reducing load on utility dispatch. Minimum 50 kW curtailable load; interval telemetry required. Season: June-September.",
         fields:[{k:"curtailableLoad",l:"Curtailable Load (kW)",t:"number",min:50,max:5000},{k:"opsContact",l:"Operations Contact (24h)",t:"text"}]},
  EVMC: {name:"EV Managed Charging", cat:"Demand Response / DER", recipe:"RECIPE_DER",
         terms:"Enroll an electric vehicle or networked charger via telematics or the charger network. Charging shifts to off-peak windows; per-session or monthly incentives. You can always override for immediate charging.",
         fields:[{k:"enrollMethod",l:"Enrollment Method",t:"select",opts:["Vehicle telematics","Networked charger"]},{k:"vehicleModel",l:"Vehicle / Charger Model",t:"text"},{k:"startDate",l:"Requested Start Date",t:"date"}]},
  EEREB:{name:"Energy-Efficiency Rebates", cat:"Energy Efficiency", recipe:"RECIPE_EE",
         terms:"Rebates for qualifying efficiency purchases (heat pumps, HVAC, insulation, appliances). Instant or mail-in; permanently lower bills. Rebate tiers per measure and efficiency rating.",
         fields:[{k:"measure",l:"Measure",t:"select",opts:["Heat pump","HVAC replacement","Insulation","Appliance","Smart controls"]},{k:"purchaseDate",l:"Purchase Date",t:"date"},{k:"invoiceAmount",l:"Invoice Amount ($)",t:"number",min:50,max:50000}]},
  AUTOPAY:{name:"AutoPay / Recurring Payment", cat:"Payment", recipe:"RECIPE_PAYMENT",
         terms:"Automatic draft of the bill amount from a bank account or tokenized card each cycle. Never miss a payment; combines with any billing program.",
         fields:[{k:"payMethod",l:"Payment Method on File",t:"select",opts:["Bank account (direct debit)","Card (tokenized)"]},{k:"startCycle",l:"Start With Bill Cycle",t:"select",opts:["Next bill","Following bill"]}]},
  PEXT: {name:"Payment Extension", cat:"Payment", recipe:"RECIPE_PAYMENT",
         terms:"A short one-time grace period past the due date for an account in good standing. Does not suppress dunning beyond the extension date.",
         fields:[{k:"extensionDays",l:"Extension",t:"select",opts:["5 days","10 days","15 days"]},{k:"reason",l:"Reason",t:"select",opts:["Income timing","Temporary hardship","Other"]}]},
  SUMBILL:{name:"Summary / Consolidated Billing", cat:"Billing Preference", recipe:"RECIPE_BILLCONFIG",
         terms:"Combines multiple accounts or meters on one statement (joint invoicing / collective bill). For customers with many premises: C&I, landlords, agencies.",
         fields:[{k:"masterAccount",l:"Master Contract Account",t:"text"},{k:"billCycle",l:"Consolidated Bill Cycle",t:"select",opts:["Monthly - 1st","Monthly - 15th"]}]},
  TPN:  {name:"Third-Party / Duplicate Bill Notification", cat:"Assistance", recipe:"RECIPE_CORR",
         terms:"Sends a copy of bills or shutoff notices to a designated third party (relative, caregiver, agency) so protected customers are not disconnected unnoticed.",
         fields:[{k:"thirdPartyName",l:"Third-Party Name",t:"text"},{k:"thirdPartyContact",l:"Third-Party Email / Phone",t:"text"},{k:"relationship",l:"Relationship",t:"select",opts:["Relative","Caregiver","Agency","Other"]}]},
  LIHEAP:{name:"LIHEAP Energy Assistance", cat:"Assistance", recipe:"RECIPE_ASSIST",
         terms:"Federal energy-bill assistance grant for income-qualified households, applied to the account balance; crisis assistance available. Agency verification follows enrollment.",
         fields:[{k:"householdSize",l:"Household Size",t:"number",min:1,max:12},{k:"incomeBracket",l:"Annual Household Income",t:"select",opts:["Below $30,000","$30,000-$45,000","$45,000-$60,000"]}]},
  PIPP: {name:"Percentage of Income Payment Plan (PIPP)", cat:"Assistance", recipe:"RECIPE_ASSIST",
         terms:"Caps the monthly bill at a fixed percentage of verified household income. Income re-verification annually; missed payments can end the cap.",
         fields:[{k:"householdIncome",l:"Monthly Household Income ($)",t:"number",min:1,max:20000},{k:"percentCap",l:"Payment Cap",t:"select",opts:["2% of income","4% of income","6% of income"]}]},
  AMP:  {name:"Arrearage Management / Forgiveness (AMP)", cat:"Assistance", recipe:"RECIPE_PAYMENT",
         terms:"Forgives a portion of past-due debt for each month the current bill is paid on time, until the arrears are cleared. Default removes remaining forgiveness.",
         fields:[{k:"commitTerm",l:"Commitment Term",t:"select",opts:["12 months","18 months"]},{k:"startDate",l:"First On-Time Bill Month",t:"date"}]},
  GREEN:{name:"Green Power / Renewable Choice", cat:"Rate / Sustainability", recipe:"RECIPE_RATE",
         terms:"Purchase 50-100% renewable energy for a small premium per kWh, applied as a product option / surcharge on the rate.",
         fields:[{k:"renewableShare",l:"Renewable Share",t:"select",opts:["50%","100%"]},{k:"startDate",l:"Requested Start Date",t:"date"}]},
  CSOL: {name:"Community Solar Subscription", cat:"Rate / Sustainability", recipe:"RECIPE_RATE",
         terms:"Subscribe to a share of a community solar array and receive bill credits for its output - solar savings without rooftop panels. Subject to array availability; low-income capacity reserved.",
         fields:[{k:"subscriptionShare",l:"Subscription Share",t:"select",opts:["1 kW block","2 kW blocks","5 kW blocks"]},{k:"startDate",l:"Requested Start Date",t:"date"}]},
  ALERTS:{name:"Usage / High-Bill Alerts", cat:"Account Management", recipe:"RECIPE_NOTIF",
         terms:"Notifies you when usage or the projected bill crosses your threshold (requires AMI meter). Avoid bill surprises; choose your channel.",
         fields:[{k:"alertType",l:"Alert Type",t:"select",opts:["Usage threshold","Projected high bill","Both"]},{k:"threshold",l:"Threshold ($ projected bill)",t:"number",min:25,max:2000},{k:"channel",l:"Channel",t:"select",opts:["SMS","Email","Push"]}]},
  INTR: {name:"Interruptible / Firm Service Level Tariff", cat:"Demand Response / Commercial", recipe:"RECIPE_RATE",
         terms:"Year-round discounted rate for large customers who contractually agree to interrupt load down to a firm service level on utility request. Penalties apply for non-performance.",
         fields:[{k:"firmServiceLevel",l:"Firm Service Level (kW)",t:"number",min:100,max:50000},{k:"contractTerm",l:"Contract Term",t:"select",opts:["12 months","36 months","60 months"]}]},
  ADR:  {name:"Automated Demand Response (OpenADR)", cat:"Demand Response / Commercial", recipe:"RECIPE_DER",
         terms:"Additional incentives for automating curtailment through a building/energy management system that receives OpenADR dispatch signals; includes rebates on control equipment.",
         fields:[{k:"bmsVendor",l:"BMS / EMS Vendor",t:"text"},{k:"openadrVersion",l:"OpenADR Version",t:"select",opts:["2.0a","2.0b","3.0"]}]}
};

/* =================== CUSTOMER PERSONAS =================== */
const CUSTOMERS = {
  "1000004711": {
    name:"Maria Gonzalez", initials:"MG", type:"Residential",
    premise:"1420 Alamo Heights Blvd, San Antonio, TX 78209",
    cas:[{id:"300000123456",div:"Electricity"},{id:"300000123457",div:"Gas"}],
    caseId:"CAS-2026-018233",
    seed(){ return {
      participating:[
        {id:"BB",   status:"Enrolled", reqDate:"2025-11-03", enrDate:"2025-11-05", ca:"300000123456"},
        {id:"EBILL",status:"Enrolled", reqDate:"2025-11-03", enrDate:"2025-11-03", ca:"300000123456"},
        {id:"DPA",  status:"Enrolled", reqDate:"2026-04-14", enrDate:"2026-04-15", ca:"300000123456",
         exitBlock:"Outstanding installment balance of $342.50 on plan 300000123456/000000000042 (EC_EXIT/007)"}
      ],
      available:["PTR","DPP","TOU","THERM","PPS","EVMC","EEREB","GREEN","CSOL","ALERTS","AUTOPAY","PREPAY","MED"],
      messages:[
        {prog:"DPA", sev:"hard", code:"EC_ELIG/012", txt:"Active installment plan already exists on contract account 300000123456."},
        {prog:"CPP", sev:"hard", code:"EC_ELIG/045", txt:"Re-enrollment waiting period active until 10 Sep 2026 (de-enrolled 10 Mar 2026)."},
        {prog:"AMP", sev:"hard", code:"EC_ELIG/071", txt:"No eligible arrears &mdash; arrearage management requires a qualifying past-due balance."},
        {prog:"CSDD",sev:"warn", code:"EC_ELIG/031", txt:"Two returned payments in last 12 months &mdash; supervisor approval required."}
      ],
      history:[
        {id:"PTR", start:"2024-06-01", end:"2025-05-31", outcome:"Completed (season credits: $86.20)"},
        {id:"CPP", start:"2025-06-01", end:"2026-03-10", outcome:"De-enrolled &mdash; customer request (waiting period applies)"}
      ]};}
  },
  "1000007322": {
    name:"Robert Chen", initials:"RC", type:"Residential",
    premise:"312 Mission Street Apt 2, San Antonio, TX 78210",
    cas:[{id:"300000129876",div:"Electricity"}],
    caseId:"CAS-2026-018377",
    seed(){ return {
      participating:[
        {id:"EBILL",status:"Enrolled", reqDate:"2024-08-10", enrDate:"2024-08-10", ca:"300000129876"}
      ],
      available:["DPA","CARE","LIHEAP","PIPP","AMP","PEXT","PREPAY","TOU","AUTOPAY"],
      messages:[
        {prog:"PTR", sev:"hard", code:"EC_ELIG/018", txt:"AMI meter communication disabled &mdash; open meter ticket MT-4471 must be resolved first."},
        {prog:"ALERTS", sev:"hard", code:"EC_ELIG/018", txt:"AMI meter communication disabled &mdash; open meter ticket MT-4471 must be resolved first."},
        {prog:"BB",  sev:"info", code:"EC_ELIG/062", txt:"Not yet eligible &mdash; requires 12 months billing history at premise (7 of 12)."},
        {prog:"CSDD",sev:"warn", code:"EC_ELIG/031", txt:"Two returned payments in last 12 months &mdash; supervisor approval required."}
      ],
      history:[
        {id:"DPA", start:"2024-02-15", end:"2024-11-20", outcome:"Completed &mdash; balance cleared on schedule"}
      ]};}
  },
  "1000001188": {
    name:"Dorothy Wilson", initials:"DW", type:"Residential",
    premise:"97 Elm Grove Lane, San Antonio, TX 78212",
    cas:[{id:"300000110234",div:"Electricity"},{id:"300000110235",div:"Gas"}],
    caseId:"CAS-2026-018395",
    seed(){ return {
      participating:[
        {id:"MED", status:"Enrolled", reqDate:"2023-05-02", enrDate:"2023-05-04", ca:"300000110234"},
        {id:"BB",  status:"Enrolled", reqDate:"2022-01-10", enrDate:"2022-01-12", ca:"300000110234"}
      ],
      available:["CSDD","EBILL","THERM","TPN","GREEN","LIHEAP"],
      messages:[
        {prog:"PREPAY", sev:"hard", code:"EC_ELIG/052", txt:"Prepayment not permitted &mdash; active Medical Alert / life-support equipment at premise."},
        {prog:"PTR",    sev:"hard", code:"EC_ELIG/044", txt:"Requires AMI meter &mdash; premise has AMR device (exchange candidate list)."},
        {prog:"CPP",    sev:"hard", code:"EC_ELIG/044", txt:"Requires AMI meter &mdash; premise has AMR device (exchange candidate list)."}
      ],
      history:[]};}
  },
  "2000003456": {
    name:"Harborview Apartments LLC", initials:"HA", type:"Commercial",
    premise:"2100 Broadway Commercial Plaza, San Antonio, TX 78215",
    cas:[{id:"300000201188",div:"Electricity"}],
    caseId:"CAS-2026-018410",
    seed(){ return {
      participating:[
        {id:"EBILL",status:"Enrolled", reqDate:"2023-03-01", enrDate:"2023-03-01", ca:"300000201188"}
      ],
      available:["TOU","CDR","ADR","INTR","SUMBILL","EEREB","GREEN","BB"],
      messages:[
        {prog:"PTR",  sev:"hard", code:"EC_ELIG/002", txt:"Restricted to residential rate class."},
        {prog:"THERM",sev:"hard", code:"EC_ELIG/002", txt:"Restricted to residential rate class."},
        {prog:"CARE", sev:"hard", code:"EC_ELIG/002", txt:"Restricted to residential rate class."},
        {prog:"MED",  sev:"hard", code:"EC_ELIG/002", txt:"Restricted to residential rate class."}
      ],
      history:[
        {id:"CDR", start:"2024-05-01", end:"2024-09-30", outcome:"Completed (curtailment credits: $2,140.00)"}
      ]};}
  }
};

/* =================== MOVE-IN SCENARIO =================== */
const MOVEIN = {
  bp:"1000009910", name:"James Carter", initials:"JC",
  moveInId:"MI-2026-00811",
  premise:"88 Riverwalk Terrace Apt 4B, San Antonio, TX 78205",
  start:"2026-08-15", caseId:"CAS-2026-018402",
  seed(){ return {
    participating:[],
    available:["PTR","CPP","DPP","TOU","THERM","CSDD","EBILL","AUTOPAY","ALERTS","GREEN","EVMC","PREPAY","CARE"],
    messages:[
      {prog:"DPA", sev:"info", code:"EC_ELIG/090", txt:"Not yet eligible &mdash; payment plans require an active contract with billing history."},
      {prog:"BB",  sev:"info", code:"EC_ELIG/062", txt:"Not yet eligible &mdash; requires 12 months billing history at premise."},
      {prog:"PPS", sev:"info", code:"EC_ELIG/093", txt:"Device installation is scheduled only after contract activation &mdash; offer at first bill review."}
    ],
    history:[]};}
};

let participating, available, messages, history, mode, currentBP="1000004711", seq=0;

/* =================== LOG / TOAST =================== */
const LOG=document.getElementById("log");
function ts(){return new Date().toLocaleTimeString("en-US",{hour12:false})}
function log(kind,msg){
  const cls=kind==="API"?"api":kind==="EVT"?"evt":kind==="ERR"?"err":"t";
  LOG.insertAdjacentHTML("afterbegin",`<div><span class="t">[${ts()}]</span> <span class="${cls}">[${kind}]</span> ${msg}</div>`);
}
function toast(msg){
  const t=document.getElementById("toast");
  t.innerHTML=msg; t.classList.add("show");
  clearTimeout(t._h); t._h=setTimeout(()=>t.classList.remove("show"),4200);
}

/* =================== RENDER =================== */
function statusChip(s){
  const map={"Enrolled":"st-enrolled","Enrollment Pending":"st-pending","Enrollment Request":"st-request",
    "Pending Approval":"st-approval","Enrollment Rejected":"st-rejected","Error - In Review":"st-error",
    "Pending De-enrollment":"st-pending","De-enrolled":"st-deenrolled","Cancelled":"st-deenrolled"};
  return `<span class="status ${map[s]||"st-request"}">${s}</span>`;
}
function cust(){ return CUSTOMERS[currentBP]; }
function render(){
  const C=cust();
  document.getElementById("crumb").innerHTML = mode==="movein"
    ? `Move-In ${MOVEIN.moveInId} &nbsp;&rsaquo;&nbsp; <b style="color:#fff">Customer Programs step</b>`
    : `Case ${C.caseId} &nbsp;&rsaquo;&nbsp; <b style="color:#fff">Enrollment Center</b>`;

  document.getElementById("ctxbar").innerHTML = mode!=="movein" ? `
    <div class="avatar">${C.initials}</div>
    <div><div class="name">${C.name} <span class="pill ${C.type==="Commercial"?"grey":"blue"}">${C.type}</span></div>
      <div class="muted">Business Partner ${currentBP} &middot; identity confirmed &#10004;</div></div>
    <div class="kv"><b>Contract Accounts</b>${C.cas.map(c=>`${c.id} <span class="pill blue">${c.div}</span>`).join("<br>")}</div>
    <div class="kv"><b>Premise</b>${C.premise}</div>
    <div class="kv"><b>Case</b>${C.caseId}</div>
    <div class="kv"><b>Context</b>Standalone work center<br><span class="muted">signed context token &#10004; (exp 300s)</span></div>`
  : `
    <div class="avatar">${MOVEIN.initials}</div>
    <div><div class="name">${MOVEIN.name} <span class="pill blue">Residential</span></div>
      <div class="muted">Business Partner ${MOVEIN.bp} (new) &middot; identity confirmed &#10004;</div></div>
    <div class="kv"><b>Move-In Document</b>${MOVEIN.moveInId} <span class="pill orange">Pre-active</span></div>
    <div class="kv"><b>Prospective Premise</b>${MOVEIN.premise}</div>
    <div class="kv"><b>Requested Start</b>${MOVEIN.start}</div>
    <div class="kv"><b>Context</b>Move-In guided process (embedded step)</div>`;

  document.getElementById("moveinBanner").classList.toggle("on", mode==="movein");
  const sel=document.getElementById("custSel");
  sel.style.display = (mode==="movein" || PAGE_MODE==="movein") ? "none":"inline-block";

  const P=document.getElementById("panel-part");
  document.getElementById("cnt-part").textContent=participating.length;
  P.innerHTML = participating.length===0 ? `<div class="empty">${mode==="movein"?"No participations yet &mdash; contract not active.":"No active program participations."}</div>` :
  `<table><tr><th>Program</th><th>Status</th><th>Requested</th><th>Enrolled</th><th></th></tr>` +
  participating.map((p,i)=>`<tr>
    <td><a class="prog" onclick="openTerms('${p.id}',true)">${CATALOG[p.id].name}</a><div class="muted">${CATALOG[p.id].cat} &middot; CA ${p.ca||"&mdash;"}</div></td>
    <td>${statusChip(p.status)}</td>
    <td>${p.reqDate||"&mdash;"}</td><td>${p.enrDate||"&mdash;"}</td>
    <td>${["Enrolled"].includes(p.status)?`<button class="btn" onclick="openDeEnroll(${i})">De-enroll</button>`:""}</td>
  </tr>`).join("")+"</table>";

  const A=document.getElementById("panel-avail");
  document.getElementById("cnt-avail").textContent=available.length;
  A.innerHTML = available.length===0 ? `<div class="empty">No further eligible programs.</div>` :
  `<table><tr><th>Program</th><th>Category</th><th></th></tr>` +
  available.map(id=>`<tr>
    <td><a class="prog" onclick="openTerms('${id}')">${CATALOG[id].name}</a></td>
    <td class="muted">${CATALOG[id].cat}</td>
    <td><button class="btn primary" onclick="openEnroll('${id}')">Enroll</button></td>
  </tr>`).join("")+"</table>" +
  `<div class="muted" style="margin-top:8px">Eligibility evaluated in one call for all programs across ${mode==="movein"?"prospective move-in data":cust().cas.length+" contract account(s)"} &middot; rule set v2026.06.2</div>`;

  const M=document.getElementById("panel-msg");
  document.getElementById("cnt-msg").textContent=messages.length;
  M.innerHTML = messages.length===0 ? `<div class="empty">No eligibility restrictions.</div>` :
  `<table><tr><th>Program</th><th>Severity</th><th>Message</th><th></th></tr>` +
  messages.map(m=>`<tr>
    <td><b>${CATALOG[m.prog].name.split("(")[0].trim()}</b></td>
    <td><span class="sev ${m.sev==="hard"?"hard":m.sev==="warn"?"warn":"info"}">${m.sev==="hard"?"HARD STOP":m.sev==="warn"?"WARNING":"NOT YET ELIGIBLE"}</span></td>
    <td>${m.txt}<div class="muted">${m.code}</div></td>
    <td>${m.sev==="warn"?`<button class="btn" onclick="openOverride('${m.prog}')">Request supervisor approval</button>`:""}</td>
  </tr>`).join("")+"</table>";

  const H=document.getElementById("panel-hist");
  document.getElementById("cnt-hist").textContent=history.length;
  H.innerHTML = history.length===0 ? `<div class="empty">No program history${mode==="movein"?" for this prospective premise":""}.</div>` :
  `<table><tr><th>Program</th><th>Start</th><th>End</th><th>Outcome</th></tr>` +
  history.map(h=>`<tr><td><b>${CATALOG[h.id].name}</b></td><td>${h.start}</td><td>${h.end}</td><td>${h.outcome}</td></tr>`).join("")+"</table>" +
  `<div class="muted" style="margin-top:8px">History-based rules (waiting periods) are applied automatically by the eligibility engine.</div>`;
}

/* =================== MODE + CUSTOMER =================== */
function applySeed(s){ participating=s.participating; available=s.available; messages=s.messages; history=s.history; }
function setMode(m){
  mode=m;
  if(PAGE_MODE==="combined"){
    document.getElementById("mode-standalone")?.classList.toggle("active",m==="standalone");
    document.getElementById("mode-movein")?.classList.toggle("active",m==="movein");
  }
  applySeed(m==="standalone"?cust().seed():MOVEIN.seed());
  const who = m==="movein"?MOVEIN.bp:currentBP;
  log("API",`GET /enrollment/v1/context/${who} &rarr; 200 (aggregate: header, participations, history) &middot; ${m==="movein"?"contextMode=MOVE_IN":"contextMode=STANDALONE"}`);
  log("API",`POST /enrollment/v1/eligibility:evaluate &rarr; 200 &middot; ${Object.keys(CATALOG).length} programs, one round trip (Z_EC_ELIGIBILITY_PRECHECK, parallel BRF+)`);
  render();
}
function setCustomer(bp){
  currentBP=bp;
  log("API",`Context switch &mdash; identified customer BP ${bp} (${CUSTOMERS[bp].name})`);
  setMode("standalone");
}

/* =================== MODAL HELPERS =================== */
const OV=document.getElementById("overlay"), MO=document.getElementById("modal");
function openModal(html){MO.innerHTML=html;OV.classList.add("open")}
function closeModal(){OV.classList.remove("open")}
OV.addEventListener("click",e=>{if(e.target===OV)closeModal()});

/* =================== TERMS =================== */
function openTerms(id,fromPart){
  const c=CATALOG[id];
  openModal(`<header>${c.name}<button class="x" onclick="closeModal()">&#10005;</button></header>
  <div class="mbody">
    <div class="muted" style="margin-bottom:8px">${c.cat} &middot; fulfillment recipe <b>${c.recipe}</b></div>
    <p style="line-height:1.6">${c.terms||"Program terms and description (see program catalog)."}</p>
  </div>
  <footer>
    <button class="btn ghost" onclick="closeModal()">Close</button>
    ${!fromPart && c.fields?`<button class="btn primary" onclick="openEnroll('${id}')">Start Enrollment &rsaquo;</button>`:""}
  </footer>`);
  log("API",`GET /enrollment/v1/programs?programId=${id} &rarr; 200 (catalog + input definitions)`);
}

/* =================== ENROLL WIZARD =================== */
function openEnroll(id){
  const c=CATALOG[id];
  const idem=crypto.randomUUID();
  openModal(`<header>Enroll &mdash; ${c.name}<button class="x" onclick="closeModal()">&#10005;</button></header>
  <div class="mbody">
    <div class="steps"><div class="on">1 &middot; Program data</div><div>2 &middot; Consent</div><div>3 &middot; Confirmation</div></div>
    ${mode==="movein"?`<div class="notice info">Pre-active mode: request will be held as <b>Enrollment Request</b> and auto-fulfilled when Move-In ${MOVEIN.moveInId} activates.</div>`:""}
    <form id="enrollForm">
    ${c.fields.map(f=>`<div class="field"><label>${f.l} <span class="req">*</span></label>
      ${f.t==="select"?`<select name="${f.k}" required><option value="">&mdash; select &mdash;</option>${f.opts.map(o=>`<option>${o}</option>`).join("")}</select>`
       :`<input name="${f.k}" type="${f.t}" required ${f.t==="number"?`min="${f.min??1}" max="${f.max??999}"`:""}>`}</div>`).join("")}
    <div class="consent"><input type="checkbox" id="consent" required style="margin-top:2px">
      <label for="consent">Customer verbally consented to the program terms, and the consent statement was read in full. Consent is recorded with the enrollment (statement ID ${id}_TERMS_V3).</label></div>
    </form>
    <div class="muted" style="margin-top:10px">Idempotency-Key: <span style="font-family:Consolas">${idem}</span></div>
  </div>
  <footer>
    <button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" onclick="submitEnroll('${id}','${idem}')">Submit Enrollment</button>
  </footer>`);
}

function submitEnroll(id,idem){
  const f=document.getElementById("enrollForm");
  if(!f.reportValidity())return;
  closeModal();
  log("API",`POST /enrollment/v1/enrollments (Idempotency-Key: ${idem.slice(0,8)}&hellip;) &mdash; server-side eligibility re-validation&hellip;`);
  setTimeout(()=>{
    log("API",`&rarr; 201 Created &middot; requestId REQ-${(1000+ ++seq)} &middot; status <b>Enrollment Request (Submitted)</b> &middot; audit &#10004;`);
    const rec={id, status:"Enrollment Request", reqDate:today(), enrDate:"", ca: mode==="movein"?"(pre-active)":cust().cas[0].id, preActive: mode==="movein"};
    participating.push(rec);
    available=available.filter(x=>x!==id);
    render();
    toast(`&#9989; <b>${CATALOG[id].name}</b> submitted &mdash; status <b>Enrollment Request</b>. Case timeline updated.`);
    log("API",`POST SC V2 /case-service/cases/${mode==="movein"?MOVEIN.caseId:cust().caseId}/timeline &rarr; 200 (status entry added)`);
    if(mode!=="movein") dispatchFulfillment(rec);
    else log("EVT",`Request parked pre-active &mdash; awaiting utilco/s4/movein/completed for ${MOVEIN.moveInId}`);
  },900);
}

function dispatchFulfillment(rec){
  setTimeout(()=>{
    rec.status="Enrollment Pending"; render();
    log("EVT",`utilco/ec/fulfillment/dispatch &rarr; IF_EC_Fulfillment_Dispatch &rarr; ${CATALOG[rec.id].recipe} (S/4HANA Utilities)`);
  },1600);
  setTimeout(()=>{
    rec.status="Enrolled"; rec.enrDate=today(); render();
    log("EVT",`utilco/s4/enrollment/confirmed &larr; S/4HANA (${CATALOG[rec.id].recipe} committed) &middot; status &rarr; <b>Enrolled</b>`);
    log("API",`POST SC V2 timeline &rarr; 200 &middot; agent notified without reopening Enrollment Center (US-12)`);
    toast(`&#127881; <b>${CATALOG[rec.id].name}</b> confirmed by S/4HANA Utilities &mdash; status <b>Enrolled</b>.`);
  },4300);
}

/* =================== SUPERVISOR OVERRIDE =================== */
function openOverride(id){
  const m=messages.find(x=>x.prog===id && x.sev==="warn");
  openModal(`<header>Supervisor Approval &mdash; ${CATALOG[id].name}<button class="x" onclick="closeModal()">&#10005;</button></header>
  <div class="mbody">
    <div class="notice warn"><b>WARNING-level eligibility failure:</b> ${m?m.txt:""} (${m?m.code:""}).
      Warning-level failures may proceed with supervisor approval. <b>Hard stops can never be overridden.</b></div>
    <div class="field"><label>Business justification <span class="req">*</span></label>
      <textarea id="ovrReason" rows="3" placeholder="e.g., Returned payments caused by bank error, documentation on file"></textarea></div>
    <div class="muted">Routed via SAP Build Process Automation &rarr; supervisor inbox. Override, approver and reason are written to the audit trail (US-11).</div>
  </div>
  <footer>
    <button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" onclick="submitOverride('${id}')">Route to Supervisor</button>
  </footer>`);
}
function submitOverride(id){
  const r=document.getElementById("ovrReason").value.trim();
  if(!r){document.getElementById("ovrReason").focus();return}
  closeModal();
  log("API",`POST /enrollment/v1/enrollments/REQ-${1000+seq+1}:requestOverride &rarr; 202 &middot; SBPA process EC_SupervisorOverride started`);
  const rec={id, status:"Pending Approval", reqDate:today(), enrDate:"", ca:cust().cas[0].id};
  participating.push(rec);
  messages=messages.filter(m=>m.prog!==id);
  render();
  toast(`&#128232; Override request for <b>${CATALOG[id].name}</b> routed to supervisor inbox&hellip;`);
  setTimeout(()=>{
    log("EVT",`SBPA: supervisor <b>K. Osei</b> APPROVED override (reason recorded) &middot; OverrideRecord written`);
    rec.status="Enrollment Request"; render();
    toast(`&#9989; Supervisor approved. <b>${CATALOG[id].name}</b> proceeding to submission.`);
    dispatchFulfillment(rec);
  },3800);
}

/* =================== DE-ENROLL =================== */
function openDeEnroll(idx){
  const p=participating[idx], c=CATALOG[p.id];
  log("API",`POST /enrollment/v1/deenrollments:validate (Z_EC_DEENROLL_CHECK) &mdash; exit rules for ${p.id}&hellip;`);
  if(p.exitBlock){
    openModal(`<header>De-enroll &mdash; ${c.name}<button class="x" onclick="closeModal()">&#10005;</button></header>
    <div class="mbody">
      <div class="notice err"><b>De-enrollment blocked by exit rule:</b><br>${p.exitBlock}</div>
      <div class="muted">Exit rules (minimum participation period, outstanding balances, pending device work) are validated in S/4HANA Utilities before any de-enrollment request is created (US-15).</div>
    </div>
    <footer><button class="btn primary" onclick="closeModal()">Understood</button></footer>`);
    log("ERR",`&rarr; 422 &middot; exit rule failed: ${p.exitBlock.split("(")[1]?.replace(")","")||"EC_EXIT"}`);
    return;
  }
  openModal(`<header>De-enroll &mdash; ${c.name}<button class="x" onclick="closeModal()">&#10005;</button></header>
  <div class="mbody">
    <div class="notice ok">Exit rules validated &mdash; no blocking conditions.</div>
    <div class="field"><label>De-enrollment reason code <span class="req">*</span></label>
      <select id="deReason"><option value="">&mdash; select &mdash;</option><option>CUST_REQUEST &mdash; Customer request</option><option>MOVE_OUT &mdash; Customer moving out</option><option>DISSATISFIED &mdash; Program dissatisfaction</option><option>COST &mdash; Cost concerns</option></select></div>
  </div>
  <footer>
    <button class="btn ghost" onclick="closeModal()">Cancel</button>
    <button class="btn primary" onclick="confirmDeEnroll(${idx})">Confirm De-enrollment</button>
  </footer>`);
}
function confirmDeEnroll(idx){
  const sel=document.getElementById("deReason");
  if(!sel.value){sel.focus();return}
  const p=participating[idx];
  closeModal();
  p.status="Pending De-enrollment"; render();
  log("API",`POST /enrollment/v1/deenrollments &rarr; 201 &middot; reason ${sel.value.split(" ")[0]} &middot; status <b>Pending De-enrollment</b>`);
  setTimeout(()=>{
    history.push({id:p.id, start:p.enrDate, end:today(), outcome:"De-enrolled &mdash; "+sel.value.split(" \u2014 ")[1]});
    participating=participating.filter((_,i)=>i!==idx);
    if(!available.includes(p.id) && CATALOG[p.id].fields) available.push(p.id);
    render();
    log("EVT",`utilco/s4/enrollment/confirmed &larr; S/4HANA &middot; de-enrollment committed &middot; history updated`);
    toast(`&#9989; <b>${CATALOG[p.id].name}</b> de-enrolled. Waiting-period rules now apply automatically.`);
  },3200);
}

/* =================== MOVE-IN EVENTS =================== */
function simulateMoveInComplete(){
  const parked=participating.filter(p=>p.preActive && p.status==="Enrollment Request");
  log("EVT",`utilco/s4/movein/completed &larr; S/4HANA &middot; ${MOVEIN.moveInId} &middot; contract activated ${MOVEIN.start}`);
  if(parked.length===0){toast("Move-In completed &mdash; no parked enrollment requests to fulfill. Enroll a program first, then fire the event.");return}
  log("API",`Eligibility re-validated on active contract for ${parked.length} parked request(s) (US-17)`);
  parked.forEach(rec=>{rec.preActive=false; rec.ca="300000123458"; dispatchFulfillment(rec);});
  toast(`&#9654; Move-In completed &mdash; ${parked.length} enrollment request(s) auto-dispatched to fulfillment.`);
}
function simulateMoveInCancel(){
  const parked=participating.filter(p=>p.preActive);
  log("EVT",`utilco/s4/movein/cancelled &larr; S/4HANA &middot; ${MOVEIN.moveInId}`);
  if(parked.length===0){toast("Move-In cancelled &mdash; no linked enrollment requests existed.");return}
  parked.forEach(p=>p.status="Cancelled");
  participating=participating.filter(p=>!p.preActive || p.status!=="Cancelled");
  parked.forEach(p=>{if(CATALOG[p.id].fields && !available.includes(p.id)) available.push(p.id)});
  render();
  log("API",`Cascade: ${parked.length} linked Enrollment Request(s) cancelled automatically &middot; audit &#10004; (US-18)`);
  toast(`&#10005; Move-In cancelled &mdash; ${parked.length} linked enrollment request(s) cascade-cancelled (US-18).`);
}

/* =================== INIT =================== */
function today(){return new Date().toISOString().slice(0,10)}

(function init(){
  const mc=document.getElementById("modeControls");
  if(PAGE_MODE==="combined"){
    mc.innerHTML=`<div class="mode-toggle">
      <button id="mode-standalone" class="active" onclick="setMode('standalone')">Work Center</button>
      <button id="mode-movein" onclick="setMode('movein')">Move-In (embedded)</button></div>`;
  } else {
    mc.innerHTML=`<span class="badge-mode">${PAGE_MODE==="movein"?"Move-In (embedded step)":"Work Center"}</span>`;
  }
  const sel=document.getElementById("custSel");
  sel.innerHTML=Object.entries(CUSTOMERS).map(([bp,c])=>
    `<option value="${bp}">${c.name} &mdash; ${c.type}${bp==="1000007322"?" (credit)":bp==="1000001188"?" (medical)":""}</option>`).join("");
  log("API",`Signed context token validated (agent A.Rivera) &mdash; no re-identification (US-01)`);
  setMode(PAGE_MODE==="movein"?"movein":"standalone");
})();
</script>
</body>
</html>
"""

OUT = r"C:\Users\jnamm\OneDrive\Desktop\enrollment-center-demo"
pages = [
    ("index.html",      "combined",   "Customer Enrollment Center — SAP Service Cloud V2 (Demo)"),
    ("workcenter.html", "standalone", "Enrollment Center — Work Center (Demo)"),
    ("movein.html",     "movein",     "Enrollment Center — Move-In Embedded (Demo)"),
]
for fname, pmode, title in pages:
    html = TEMPLATE.replace("__PAGEMODE__", pmode).replace("__TITLE__", title)
    with io.open(os.path.join(OUT, fname), "w", encoding="utf-8") as f:
        f.write(html)
    print("wrote", fname, len(html), "bytes")
