const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, LevelFormat,
  AlignmentType, BorderStyle, TabStopType, TabStopPosition,
} = require("docx");

// ---- Build mode ----
// `node build_resume.js --public` omits the phone number (safe for public GitHub).
// Default build includes it (for direct applications), read from the PHONE env var:
//   PHONE="+55 xx xxxxx-xxxx" node build_resume.js
const PUBLIC = process.argv.includes("--public");
const PHONE = process.env.PHONE || "";

// ---- Palette ----
const TEAL = "0E6E6A";      // petrol teal accent
const CHARCOAL = "212121";  // primary text
const GRAY = "5A5A5A";      // secondary text

const FONT = "Calibri";

// ---- Helpers ----
const t = (text, opts = {}) => new TextRun({ text, font: FONT, color: CHARCOAL, size: 20, ...opts });

const sectionHeading = (text) =>
  new Paragraph({
    spacing: { before: 200, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 2 } },
    children: [t(text.toUpperCase(), { bold: true, size: 22, color: TEAL, characterSpacing: 20 })],
  });

const roleLine = (role, dates) =>
  new Paragraph({
    spacing: { before: 140, after: 20 },
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    children: [
      t(role, { bold: true, size: 21 }),
      t("\t" + dates, { size: 19, color: GRAY }),
    ],
  });

const subLine = (text) =>
  new Paragraph({
    spacing: { after: 40 },
    children: [t(text, { italics: true, size: 19, color: GRAY })],
  });

const bullet = (children, opts = {}) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 40 },
    ...opts,
    children,
  });

const b = (text) => t(text, { bold: true });

// Client/section sub-heading that never gets orphaned at a page bottom
const clientHeading = (text, dates) =>
  new Paragraph({
    spacing: { before: 100, after: 20 },
    keepNext: true,
    keepLines: true,
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    children: dates
      ? [t(text, { bold: true, size: 20, color: TEAL }), t("\t" + dates, { size: 18, color: GRAY })]
      : [t(text, { bold: true, size: 20, color: TEAL })],
  });

const link = (text, url) =>
  new ExternalHyperlink({
    link: url,
    children: [new TextRun({ text, font: FONT, size: 20, color: TEAL, underline: {} })],
  });

// ---- Document ----
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20, color: CHARCOAL } } } },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 320, hanging: 200 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 800, right: 900, bottom: 800, left: 900 },
      },
    },
    children: [
      // ================= HEADER =================
      new Paragraph({
        spacing: { after: 40 },
        children: [t("FELIPE SANTOS", { bold: true, size: 40, color: CHARCOAL, characterSpacing: 30 })],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [t("Physical AI | Edge AI | IoT Platforms | Solutions Architect | Technical Lead | M.Sc. in AIoT & Cybersecurity", { bold: true, size: 22, color: TEAL })],
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          t("Belo Horizonte, Brazil (remote-first, open to global mobility & travel)   |   ", { size: 19, color: GRAY }),
          ...(!PUBLIC && PHONE ? [t(`${PHONE}   |   `, { size: 19, color: GRAY })] : []),
          link("felipeaugustodosantos@gmail.com", "mailto:felipeaugustodosantos@gmail.com"),
        ],
      }),
      new Paragraph({
        spacing: { after: 40 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4 } },
        children: [
          link("linkedin.com/in/feli-santos", "https://www.linkedin.com/in/feli-santos/"),
          t("   |   ", { size: 19, color: GRAY }),
          link("github.com/feli-santos", "https://github.com/feli-santos"),
        ],
      }),

      // ================= SUMMARY =================
      sectionHeading("Executive Summary"),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          t("Senior technical leader with 7+ years architecting intelligent systems that connect the physical and digital worlds — "),
          b("Physical AI, Edge AI, AIoT, computer vision, robotics, and agentic AI platforms with LLM reasoning cores"),
          t(". Practice Lead for Edge AI & Platforms within Globant's Physical AI Studio and member of the studio management team. Proven record leading cross-functional teams of 5+ engineers across "),
          b("AWS, Azure, and GCP"),
          t(", delivering mission-critical platforms for global brands (LATAM Airlines, NBA/LA Clippers, Universal, Disney). Trusted client-facing solutions architect driving "),
          b("presales wins, PoCs, MVPs and multi-project account growth"),
          t(". AI-native engineering practitioner: "),
          b("agentic coding workflows, spec-driven development, context engineering, and AI agent harness design"),
          t(" for record-speed delivery. Mentor, technical interviewer (gatekeeper), and multilingual communicator (EN/PT/ES)."),
        ],
      }),

      // ================= CORE COMPETENCIES =================
      sectionHeading("Core Competencies"),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          b("AI & ML:  "),
          t("Physical AI · Edge AI · Agentic AI & Autonomous Systems · LLM Reasoning & Multi-Agent Orchestration · GenAI Platform Architecture · Computer Vision (YOLO, OpenCV, NVIDIA DeepStream, TensorRT, Triton, OpenVINO, custom model training) · MLOps (SageMaker, edge-to-cloud pipelines) · PyTorch · TensorFlow · Sensor Fusion · Reinforcement Learning · Real-Time Inference Optimization"),
        ],
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          b("IoT & Edge:  "),
          t("AIoT Platforms · IoT Device Management at Scale (20,000+ devices) · Azure IoT Hub/Edge · AWS IoT Core · OTA Updates (Balena, Mender) · Zero-Touch Provisioning · MQTT/Mosquitto · BLE Beacons & Gateways · NVIDIA Jetson · Raspberry Pi · Embedded Linux · Robotics · Digital Twins · Industrial Edge Computing · IoT Cybersecurity"),
        ],
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          b("Cloud & Platform:  "),
          t("Multi-Cloud (AWS, Azure, GCP) · CloudOps · Terraform/IaC · Docker · Kubernetes · CI/CD · Microservices & Event-Driven Architecture · Serverless (Cloud Run, Lambda, ECS) · High-Throughput Data Ingestion (Pub/Sub, Dataflow, Kinesis Firehose) · Go · Python · Bash · SQL/NoSQL (Firestore, DynamoDB)"),
        ],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          b("Leadership & AI-Native Delivery:  "),
          t("Solutions Architecture · Presales & Technical Storytelling · Team Leadership · Stakeholder Management · Mentoring · Technical Hiring · Agentic Coding (Claude Code, Codex, OpenCode) · Spec-Driven Development · Context & Token Engineering · Agent Harness Design"),
        ],
      }),

      // ================= EXPERIENCE =================
      sectionHeading("Professional Experience"),

      roleLine("GLOBANT — Technology Consultant | Tech Lead · Solutions Architect · Practice Lead", "May 2022 – Present"),
      subLine("Physical AI Studio — Edge AI, AIoT, computer vision, and robotics solutions for aviation, sports & entertainment, smart venues, utilities, and manufacturing."),

      clientHeading("Studio & Practice Leadership"),
      bullet([t("Practice Lead for "), b("Edge AI and Platforms"), t(" verticals; contributor to Robotics and Embedded practices within the Physical AI Studio.")]),
      bullet([t("Member of the studio management team: drive hackathons, tech sessions, internal R&D prototypes, and staffing strategy for studio projects.")]),
      bullet([t("Tech Lead of internal product initiatives: "), b("RobOps platform and intelligent edge devices for human-robot interaction"), t(" (Edge AI + robotics prototypes).")]),
      bullet([t("Solutions Architect in presales: technical pitches, solution design, PoCs, MVPs, and live demos contributing to "), b("multiple six-figure USD project wins"), t(", plus account extensions and upsells.")]),
      bullet([t("Mentor to multiple engineers and hiring gatekeeper conducting technical interviews for staffing candidates.")]),

      clientHeading("LATAM Airlines — Lead Architect & Engineer, Agentic AI Operations Platform", "2026 – Present"),
      bullet([t("Principal architect of an "), b("AI platform with an LLM reasoning core"), t(" for real-time monitoring, management, and control of airport turnaround operations — anomaly detection, operational risk identification, and "), b("autonomous decision-making"), t(".")]),
      bullet([t("Designed event-driven GCP architecture (Compute Engine, Pub/Sub, Cloud Run APIs, Dataflow, Firestore) in "), b("Go and Python"), t(" with "), b("multi-source sensor fusion"), t(": computer vision vendors, airline subsystems, and a "), b("proprietary BLE beacon/gateway solution"), t(" for people and asset tracking.")]),

      clientHeading("Universal Parks — Jurassic Park Ride | Tech Lead, Architect & Principal Engineer", "2025 – 2026"),
      bullet([t("Architected a "), b("real-time computer vision safety and operational-efficiency platform at the edge"), t(": left-object detection, people counting, restricted-zone monitoring, and ride dispatch metrics, integrated with operator alert panels.")]),
      bullet([t("Achieved "), b("0.95+ F1-score with <500 ms end-to-end latency"), t(" using a custom-trained YOLO model on NVIDIA DeepStream with optimized edge inference (industrial PCs, MQTT/Mosquitto, Go, AWS, Docker, CI/CD); led a team of 5 engineers with on-site demos and client-facing delivery.")]),

      clientHeading("BBVA Argentina — Innovation Consultant (Internal Hackathon)", "2025"),
      bullet([t("Advised the team on an "), b("eye-tracking applied-intelligence and cybersecurity"), t(" solution for the bank's mobile application.")]),

      clientHeading("Compass (Cosan) — Tech Lead, Smart Metering AIoT Platform", "2024 – 2025"),
      bullet([t("Led a team of 5 building an IoT platform to integrate and remotely manage "), b("millions of smart meters"), t(": automated device provisioning, device management, security, and OTA updates.")]),
      bullet([t("Engineered "), b("high-throughput IoT data ingestion and preprocessing"), t(" on AWS (IoT Core, Kinesis Firehose, Lambda, ECS, DynamoDB, S3, MQTT) with Terraform IaC in Go/Python; integrated with the data organization for advanced analytics and data intelligence.")]),

      clientHeading("LA Clippers (NBA) — Intuit Dome | Senior Engineer promoted to Tech Lead", "2023 – 2024"),
      bullet([t("Built and led delivery of an enterprise "), b("device management platform for 20,000+ devices"), t(" at the Intuit Dome smart venue — remote control, fleet monitoring, OTA updates, zero-touch provisioning, cybersecurity, and automated deployments.")]),
      bullet([t("Led the IoT Edge and IoT Platform engineering teams on the full "), b("Azure IoT ecosystem"), t(" (IoT Hub, IoT Edge) with Terraform, Docker, Python, and CloudOps; promoted from engineer to tech lead during the engagement; on-site delivery and client strategy participation.")]),

      clientHeading("HME — Edge-to-Cloud MLOps Platform Engineer", "2023"),
      bullet([t("Developed an end-to-end "), b("edge-to-cloud MLOps platform"), t(" for computer vision models on the NVIDIA stack (DeepStream, YOLO) with Python, Docker, and Bash.")]),

      clientHeading("Disney Parks — MLOps Platform Engineer", "2022 – 2023"),
      bullet([t("Built the end-to-end "), b("MLOps platform"), t(" (deployment, monitoring, data consumption, training, and inference) for a computer vision solution improving ride operational efficiency — AWS SageMaker, DynamoDB, Terraform, Python, CI/CD.")]),

      // ---- Earlier experience ----
      roleLine("BRPHOTONICS — Automation Analyst promoted to Automation & Control Engineer", "Apr 2020 – Apr 2022"),
      subLine("R&D and manufacturing of optoelectronic devices (lasers, amplifiers) for optical communications."),
      bullet([t("Designed and fully automated manufacturing, testing, calibration, and quality processes for lasers, amplifiers, and other optoelectronic devices; architected pilot production lines from R&D approval to manufacturing scale.")]),
      bullet([t("Led junior engineers, interns, and technicians across automation, test, and quality engineering; time-series and statistical data analysis for product performance validation.")]),

      roleLine("IDEA! ELECTRONIC SYSTEMS — Product Engineering Intern", "Jul 2019 – Mar 2020"),
      bullet([t("Automated calibration, testing, and quality processes for "), b("5G infrastructure optoelectronic devices"), t(".")]),

      // ================= EDUCATION =================
      sectionHeading("Education & Research"),

      roleLine("M.Sc. in Computer Engineering — UNICAMP (University of Campinas), Brazil", "2024 – 2026"),
      bullet([t("Thesis: "), b("Reinforcement Learning applied to cybersecurity in IoT device networks"), t(" — designed and built an RL agent training framework for autonomous cyber-attack response using the military Kill Chain taxonomy.")]),
      bullet([t("Paper under review at Elsevier "), t("Internet of Things", { italics: true }), t(": "), link("view paper (SSRN)", "https://papers.ssrn.com/sol3/papers.cfm?abstract_id=7370791")]),
      bullet([t("Collaborated on graduate research projects in digital twins, advanced data analytics, robotics, and AI.")]),

      roleLine("B.Sc. in Electrical Engineering (Control & Automation) — UFPB, Brazil", "2014 – 2020"),
      bullet([t("CNPq/PIBIC research fellow (2016–2019): ultrasonic anemometry for wind velocity measurement; member of the Applied Computational Intelligence group; "), b("4 peer-reviewed publications"), t(" (IEEE, IOP): "), link("IEEE (2018)", "https://ieeexplore.ieee.org/document/8546717"), t("  ·  "), link("IEEE (2019)", "https://ieeexplore.ieee.org/document/8868403"), t("  ·  "), link("IOP J. Phys. (2018)", "https://iopscience.iop.org/article/10.1088/1742-6596/1044/1/012045")]),
      bullet([t("Science Without Borders scholar (CAPES): exchange year in Electrical & Computer Engineering at the "), b("University of Wisconsin–Milwaukee, USA"), t(" (2015–2016); Assistant Researcher at "), b("Mercer University, USA"), t(" (2016) — wireless power transfer via resonant inverters.")]),

      new Paragraph({
        spacing: { before: 60, after: 0 },
        children: [
          b("Languages:  "),
          t("Portuguese (native)  ·  English (full professional)  ·  Spanish (full professional)"),
        ],
      }),
    ],
  }],
});

const outFile = PUBLIC
  ? "/Users/felipe.santos/Documents/resume/Felipe_Santos_Resume_Public.docx"
  : "/Users/felipe.santos/Documents/resume/Felipe_Santos_Resume.docx";

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outFile, buffer);
  console.log(`DOCX written: ${outFile}${PUBLIC ? " (no phone)" : ""}`);
});
