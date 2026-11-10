# Changelog

All notable changes to the Resident Research Progression Tracker will be documented in this file.

## [v1.2.0] - 2026-11-10
### Added
- **Complete Faculty Roster (33 Advisors)**: Fully updated and categorized across 7 ophthalmology subspecialties:
  - Cornea (5): พญ.นวลจิรา, พญ.กนกทิพย์, พญ.เบญจมา, นพ.วีรภัทร, พญ.ชิดชนก
  - Glaucoma (5): พญ.ดวงดาว, พญ.กุลวรรณ, พญ.อรอร, พญ.สุพัตรา, นพ.ปิติพงศ์
  - Neuro-Ophthalmology (3): พญ.ปิยวดี, นพ.ปรเมศวร์, นพ.จิเมธัส
  - Ophthalmic Plastic & Reconstructive Surgery / OPRC (4): พญ.สุณิสา, พญ.มิ่งขวัญ, พญ.อรวีณัฏฐ์, นพ.สรรเสริญ
  - Pediatric Ophthalmology & Strabismus (2): พญ.อาภา, พญ.จิราภา
  - Retina & Vitreous (11): พญ.สุดาวดี, พญ.วงศ์ศิริ, นพ.วรภัทร, พญ.อุมาภรณ์, พญ.พิริญา, นพ.เอกชัย, พญ.กาญจน์ธิดา, นพ.วรินทร์, พญ.นวลพร, พญ.ทิฆัมพร, พญ.ปาณิศา
  - Uveitis & Ocular Immunology (3): นพ.โยธิน, นพ.นพคุณ, พญ.ณิศวรรณ
- **Clean & Modern UI Redesign**:
  - macOS / Linear-style Segmented Pill Control for the 7 dashboard tabs.
  - GSAP kinetic number counting for high-impact KPI summary cards.
  - Strict typographic hierarchy using Plus Jakarta Sans for numbers and Sarabun/Prompt for Thai copy.
  - Tabular numeric alignment (`tabular-nums`) across all data tables and metrics.
- **Production Cloud Deployment**:
  - Live Google Apps Script deployment on `research.metta@gmail.com`.
  - Continuous 2-way sync with Google Sheets central database.
  - GitHub Pages static web app hosting capability.

## [v1.1.0] - 2026-11-05
### Added
- Multi-cohort Gantt & Interactive Kanban board.
- Automated At-Risk Early Warning Radar.
- Resident Self-Service Portal with Google Drive attachment linkage.
- Role-based Admin PIN authorization gate (Factory Default: `1234`).
- Google Sheets 2-way synchronization via Google Apps Script (GAS).

## [v1.0.0] - 2026-11-02
### Initial Release
- Core 6-stage residency research curriculum architecture.
- Standalone SPA prototype and Excel database generator.
