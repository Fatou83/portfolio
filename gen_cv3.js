const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  ExternalHyperlink, TabStopType, TabStopPosition, BorderStyle, HeadingLevel
} = require("docx");

function sectionTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
    children: [new TextRun({ text, bold: true, color: "2E75B6" })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 30 },
    children: [new TextRun({ text, size: 20 })],
  });
}

function jobHeader(title, dates) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { before: 100, after: 0 },
    children: [
      new TextRun({ text: title, bold: true, size: 21 }),
      new TextRun({ text: "\t" + dates, bold: true, size: 21 }),
    ],
  });
}

function companyLine(company, location) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({ text: company, italics: true, size: 20 }),
      new TextRun({ text: location ? (", " + location) : "", italics: true, size: 20 }),
    ],
  });
}

function bodyParagraph(text) {
  return new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: text, size: 20 })] });
}

const children = [];

// Header
children.push(new Paragraph({
  spacing: { after: 20 },
  children: [new TextRun({ text: "FATOU MANE NDIAYE", bold: true, size: 32 })],
}));
children.push(new Paragraph({
  spacing: { after: 80 },
  children: [new TextRun({ text: "Consultante informatique et Développeuse", size: 20, color: "2E75B6", bold: true })],
}));
children.push(new Paragraph({
  spacing: { after: 40 },
  children: [new TextRun({ text: "Sacrée Cœur, Dakar  |  +221-77-748-98-85  |  fatoumanendiaye83@gmail.com", size: 18 })],
}));
children.push(new Paragraph({
  spacing: { after: 120 },
  children: [
    new ExternalHyperlink({ link: "https://www.linkedin.com/in/fatou-man%C3%A9-ndiaye-2a87b4167/", children: [new TextRun({ text: "LinkedIn", style: "Hyperlink", size: 18 })] }),
    new TextRun({ text: "  |  ", size: 18 }),
    new ExternalHyperlink({ link: "https://fatou83.github.io/portfolio/", children: [new TextRun({ text: "https://fatou83.github.io/portfolio/", style: "Hyperlink", size: 18 })] }),
    new TextRun({ text: "  |  Permis de conduire : B", size: 18 }),
  ],
}));

// PROFIL
children.push(sectionTitle("PROFIL"));
children.push(bodyParagraph(
  "Développeuse web et mobile passionnée, avec une solide expertise en développement front-end et back-end et en conception d'applications mobiles hybrides (Flutter, Ionic, Angular, Spring, Node.js, Keycloak). Capable de transformer des idées en produits numériques innovants, avec une approche centrée sur l'utilisateur. Sensibilisée à la conception de dispositifs de collecte et de suivi-évaluation de données, apportant une vision transversale entre développement technique et exploitation des données. Dynamique, adaptable et orientée résultats."
));

// EXPERIENCE
children.push(sectionTitle("EXPÉRIENCE PROFESSIONNELLE"));

children.push(jobHeader("Consultante informatique et Développeuse", "06/2025 – Présent"));
children.push(companyLine("CINERI", "Dakar, Sénégal"));
[
  "Mise en place d'une authentification centralisée et d'un plugin OTP pour l'authentification multi-facteur (MFA).",
  "Développement d'un thème personnalisé pour Keycloak.",
  "Cheffe de projet de l'application de gestion des inscriptions des étudiants (Spring Boot/Angular).",
  "Conception du module de gestion d'accréditation pour l'application de gestion des orientations (Spring Boot/Angular).",
  "Conception d'une application de gestion de délivrance de diplômes pour les bacheliers (Spring Boot/Angular).",
].forEach(function(t) { children.push(bullet(t)); });

children.push(jobHeader("Consultante informatique et Développeuse", "11/2024 – 02/2025"));
children.push(companyLine("TechCo", "Dakar, Sénégal"));
[
  "Création d'une application de gestion des bulletins de salaire.",
  "Intégration de l'API Postmark (e-mails) et des APIs SMS/WhatsApp pour notifier les utilisateurs.",
].forEach(function(t) { children.push(bullet(t)); });

children.push(jobHeader("Consultante informatique et Développeuse", "06/2024 – 11/2024"));
children.push(companyLine("Consi", "Remote, Sénégal"));
[
  "Développement d'une application mobile (Flutter) pour les patients, avec intégration des APIs backend.",
  "Refonte du site web sur WordPress.",
  "Intégration des APIs SMS/WhatsApp et OTP pour les notifications, et de l'API PayPal (Node.js/Flutter) pour la gestion d'annonces.",
].forEach(function(t) { children.push(bullet(t)); });

children.push(jobHeader("Développeuse Web et Mobile", "02/2022 – 06/2024"));
children.push(companyLine("Saou Transport", "Dakar"));
[
  "Développement et refonte d'une application mobile (Ionic/Flutter) ; compilation iOS via Xcode, plugins Capacitor, déploiement sur l'App Store.",
  "Ajout du module de gestion des colis (web) avec Spring/Angular et déploiement avec Docker.",
  "Analyse des besoins clients, conception technique, tests, maintenance et optimisation continue.",
  "Plateforme de mise en relation patients-professionnels de santé : recherche/géolocalisation de médecins, prise de rendez-vous en ligne, gestion des comptes et notifications.",
].forEach(function(t) { children.push(bullet(t)); });

children.push(new Paragraph({
  spacing: { before: 100, after: 0 },
  children: [new TextRun({ text: "Développeuse web et mobile d'une application de transport de type IoT", bold: true, size: 21 })],
}));
children.push(new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
  spacing: { after: 40 },
  children: [
    new TextRun({ text: "Sn software, Dakar, Sénégal", italics: true, size: 20 }),
    new TextRun({ text: "\t06/2021 – 05/2022", bold: true, size: 21 }),
  ],
}));
[
  "Création d'interfaces et intégration d'API (Angular/Ionic) ; choix de l'architecture microservices et développement du gateway.",
  "Conception d'une application mobile (Flutter) et module de gestion des colis (Spring/Angular).",
  "Intégration de l'API d'une société partenaire pour l'achat de billets (Java Spring Boot/Angular).",
  "Tableau de bord statistiques et géolocalisation des bus en temps réel ; application pour compagnie aérienne (Odoo).",
].forEach(function(t) { children.push(bullet(t)); });

// FORMATION
children.push(sectionTitle("FORMATION"));
function eduLine(school, degree, city, dates) {
  children.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { after: 20 },
    children: [
      new TextRun({ text: school + " — " + degree + ", " + city, size: 20 }),
      new TextRun({ text: "\t" + dates, bold: true, size: 20 }),
    ],
  }));
}
eduLine("INSTITUT SUP'INFO", "Master, Dév. Web et Mobile", "Dakar", "2019 – 2021");
eduLine("INSTITUT SUP'INFO", "Licence 3, Dév. Web et Mobile", "Dakar", "2017 – 2019");
children.push(bullet("Projet de fin d'études : application de collecte de données pour agents de terrain (Angular, Node.js), avec module d'analyse et de prédiction en Python (NumPy)."));
eduLine("ISI", "Licence 3, Réseaux informatiques", "Dakar", "2016 – 2018");
eduLine("UCAD", "Licence 1, Science physique", "Dakar", "2013 – 2015");
eduLine("LCND", "Bac S2", "Fatick", "2012 – 2013");

// CONNAISSANCES
children.push(sectionTitle("CONNAISSANCES"));
children.push(bodyParagraph(
  "Java, Angular, Spring, Hibernate, JPA, Ionic, Flutter, Keycloak, JavaScript, Node.js, Python (NumPy), MySQL, PostgreSQL, Maven, JHipster, Linux, Windows, macOS, API REST, GitHub, Docker, App Store, Play Store, CRM, CMS, WordPress, Odoo, Drupal 8/9, HTML, XML, API Postmark, DocuSign, Conception et administration de site, Conception de base de données, SQLite, Oracle."
));
children.push(bodyParagraph("Qualités : orientée résultats, adaptable, dynamique."));

// LANGUES & CERTIFICATS
children.push(sectionTitle("LANGUES"));
children.push(bodyParagraph("Français — courant  |  Anglais — intermédiaire (B1)  |  Wolof — courant"));

children.push(sectionTitle("CERTIFICATS"));
[
  "Permis de conduire : B",
  "Concevoir un dispositif de suivi-évaluation — Groupe AFD (Agence Française de Développement)",
].forEach(function(t) { children.push(bullet(t)); });

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 640, hanging: 320 } } } }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 900, right: 1000, bottom: 900, left: 1000 },
      },
    },
    children: children,
  }],
});

Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync("CV_Fatou_Mane_Ndiaye_2pages.docx", buffer);
  console.log("done");
});
