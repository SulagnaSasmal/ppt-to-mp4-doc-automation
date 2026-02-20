PPT-to-MP4 Utility
Author: Sulagna
Status: MVP in progress
Document Type: Product / Technical Specification
Last Updated: 20/02/2026

Overview
The PPT-to-MP4 Utility is a lightweight, API-driven tool designed to convert PowerPoint presentations into MP4 video format. This utility allows for the quick creation of explainer videos, demos, and walkthroughs, especially for product presentations, documentation, and training purposes, without the need for professional video editing software.

This initiative was proposed and conceptualized to bridge the gap between static content (PPT) and scalable video assets, while keeping the workflow simple, automatable, and API-driven for future integrations.

What this is
The PPT-to-MP4 Utility is a phased, API-driven capability that converts PowerPoint presentations into MP4 videos. It enables rapid creation of explainer, demo, and walkthrough videos without manual screen recording or video-editing tools.

Why it matters
• Reduces time and effort to create and update videos
• Enables reuse of existing PPT assets at scale
• Lays the foundation for automated, AI-assisted video enablement

What’s delivered so far (Phase 1)
• A working MVP that converts static PPT slides into MP4
• Validated feasibility through a live demo to stakeholders

What’s in progress (Phase 2)
• Hardening and UX refinement for broader internal usage
• Configurability and preview-first author workflow
• Operational telemetry for performance tuning

Latest Implemented Enhancements (20 Feb 2026)
• Preview Before Convert
  o Users can preview extracted slide notes before starting a job
  o Prevents low-quality outputs due to missing/incorrect notes
• Configurable Conversion Settings
  o Voice name, speaking rate, resolution, FPS, and quality are configurable from UI
  o Settings are persisted with each job for traceability
• Queue / History Experience
  o Added history page with recent jobs, status, timestamps, logs, and quick download actions
  o Added history API for future dashboards
• Telemetry-lite
  o Captures stage-level timings (PowerPoint export, TTS, FFmpeg mux, total)
  o Exposed in job status and UI summary for diagnostics and optimization
• Packaging and Dependency Check
  o Added one-command startup script
  o Added preflight checks for Python, Azure env, FFmpeg/ffprobe, and PowerPoint COM
• UI Improvements
  o Drag-and-drop upload on the front page
  o Live log streaming panel
  o Collapsible preview section for cleaner page layout

What this unlocks next
• Scalable video generation across teams
• Future enhancements such as narration, branding, and CMS integration

Ownership & Status
• Proposed, conceptualized, and initially implemented as a POC by Sulagna
• Currently evolving through a structured MVP roadmap

Problem Statement
Today, creating short product or feature videos typically requires:
• Manual screen recording
• Dedicated video-editing tools
• High effort for minor updates

This leads to:
• Slow turnaround time
• Poor scalability
• Duplication of effort for small content changes

There is a need for a repeatable, low-effort mechanism to transform existing presentation assets into video formats that can be reused across portals, demos, and internal/external communication.

Goals & Non-Goals
Goals
• Enable automated conversion of PPT files to MP4
• Support incremental enhancement from static to animated output
• Provide an API-first design for future platform integration
• Minimize manual intervention

Non-Goals (Initial Phases)
• Advanced video editing (cuts, overlays, transitions beyond PPT animations)
• Voice-over synthesis (to be evaluated in future phases)
• Real-time video generation at scale (initially async)

Target Users & Use Cases
Target Users
• Product Managers
• Technical Writers
• Enablement & Presales teams
• Internal SMEs

Key Use Cases
• Product walkthrough videos
• Feature announcement clips
• Documentation explainers
• Internal demos and training content

Phased MVP Approach
The utility is being developed using a phase-by-phase MVP strategy, allowing early validation while progressively increasing capability.

Phase 1: Static PPT-to-MP4 (Completed / Baseline MVP)
Description
Phase 1 focuses on establishing the core feasibility of converting a PowerPoint presentation into a video file with minimal logic and dependencies.
Each slide is rendered sequentially and exported as a static frame into an MP4 video.

Capabilities
• Input: PPT/PPTX file
• Output: MP4 video
• Static slide rendering (no animations)
• Fixed duration per slide

Architecture (High-Level)
• PPT ingestion module
• Slide rendering engine
• Video compilation step

Outcomes
• Validated technical feasibility
• Demonstrated working POC
• Successfully demoed to stakeholders

Limitations
• No slide animations
• No narration or audio
• Uniform timing across slides

Phase 2: Animated PPT-to-MP4 with Hosted API (Completed)
Description
Phase 2 enhances the utility to support PPT-native animations and exposes the conversion capability via a hosted API, enabling programmatic access and future integrations.
This phase moves the solution from a local/static utility to a service-oriented capability.

Key Enhancements
• Native PPT animations and transitions preserved
• Speaker notes converted to AI-generated voice-over (Azure TTS)
• Job-based asynchronous processing
• API-first architecture with HTML UI overlay
• Progress tracking, logging, and downloadable outputs

Proposed API Design (High-Level)
Endpoint:
POST /convert
GET /status/{job_id}
GET /download/{job_id}
GET /logs/{job_id}

API Design (Implemented)
Endpoint        Purpose
POST /convert   Upload PPT and start conversion
GET /status/{job_id}    Check job progress
GET /download/{job_id}  Download final MP4
GET /logs/{job_id}      Retrieve processing logs
POST /preview-notes      Preview extracted slide notes + selected settings
GET /history             HTML queue/history page for recent jobs
GET /api/history         JSON history endpoint (recent jobs)

Processing Model
• PPT uploaded via API/UI
• Job ID returned immediately
• Backend processes animations + audio
• UI polls status and updates progress bar
• MP4 and logs available on completion

Request Parameters:
• pptFile (binary / URL) - not configurable yet
• outputResolution (e.g., 720p, 1080p)
• slideTiming (optional)
• animationEnabled (true/false) - always ON

Response:
• Job ID
• Status endpoint for progress
• Download URL for MP4
(Exact schema to be finalized)

Architecture (Conceptual)
HTML UI (Internal Users)
  └── FastAPI (Job orchestration, status, logs)
        └── PowerPoint COM (Animation rendering)
        └── Azure TTS (Voice synthesis)
        └── FFmpeg (Audio/video mux)

┌──────────────────────────────┐
│  HTML UI (for writers)       │  ← browser
│  - Upload PPT                │
│  - Progress bar              │
│  - Status text               │
│  - Download buttons          │
└──────────────┬───────────────┘
               │ HTTP (fetch / AJAX)
┌──────────────▼───────────────┐
│  FastAPI (API layer)         │  ← core logic
│  POST /convert               │
│  GET  /status/{job_id}       │
│  GET  /download/{job_id}     │
│  GET  /logs/{job_id}         │
└──────────────┬───────────────┘
               │ Python calls
┌──────────────▼───────────────┐
│  Pipeline (engine)           │
│  - PowerPoint COM            │
│  - Azure TTS                 │
│  - FFmpeg                    │
└──────────────────────────────┘

Success Criteria
• Accurate rendering of PPT animations
• Audio-video synchronization within an acceptable tolerance
• Stable MP4 output quality
• Predictable job execution and status reporting

Phase 3: Platformization, Hosting & Advanced Enhancements (Current)
Phase 3 focuses on making the capability usable by other writers and teams by introducing centralized hosting on Windows, basic access control, and operational stability — while continuing to evolve the experience and output quality.
This phase transitions the solution from a developer-owned service to a shared internal capability, without yet positioning it as a production-grade platform.

Key Enhancement Areas
Hosting & Accessibility (Critical Enabler)
This is mandatory for adoption but not final for pilot
• Centralized hosting on a Windows-based environment
  o Required due to PowerPoint COM dependency
  o Options:
     Windows VM
     Windows-based Azure App Service (if feasible)
• Single internal URL
  o Enables usage by other writers without local setup
  o Removes dependency on developer machines
• Always-on service
  o Background worker process for PPT rendering
  o API remains available during long-running jobs

⚠️ Note:
This hosting model is intended for pilot and controlled internal usage, not large-scale production deployment.

Experience & Output Enhancements
• Configurable narration behavior
  o Enable or disable voice-over
  o Narration source:
     Speaker notes (default)
     Slide comments (optional)
• Branding support
  o Standard intro/outro slides
  o Consistent look-and-feel across generated videos
• Output flexibility (optional)
  o MP4 (default)
  o Future: WebM or short demo clips

Platform & Integration Readiness
• HTML-based UI for non-technical users
  o Upload PPT
  o View progress and logs
  o Download final MP4
• API-first design retained
  o UI consumes the same APIs exposed for automation
• Future portal integration
  o Documentation portals
  o Enablement or training workflows

Scale, Reliability & Governance (Incremental)
• Controlled concurrency
  o One job per user or limited parallelism
• Basic job lifecycle management
  o Queued → Processing → Completed / Failed
• Temporary storage management
  o Automatic cleanup of PPTs, audio, and intermediate files
• Error visibility
  o Logs exposed via UI for troubleshooting

Security & Access Considerations
• Authentication for API access
• File validation:
  o Supported PPT formats and size limits
• Automatic cleanup of temporary files and logs
• Auditability for generated assets
(Details to be defined with platform/security teams)

Metrics & Validation
Phase 3 prioritization will be guided by measurable outcomes:
• Productivity impact
  o Time saved vs. manual PPT-to-video creation
• Adoption
  o Number of active users and teams
  o Frequency of video generation
• Quality & usability
  o Feedback from enablement, demos, and training use cases
• Operational stability
  o Job success rate
  o Average processing time per presentation

Open Items & Placeholders
• <PLACEHOLDER: Final API schema>
• <PLACEHOLDER: Hosting environment>
• <PLACEHOLDER: Supported PPT features>
• <PLACEHOLDER: Performance benchmarks>

Summary
The PPT-to-MP4 Utility demonstrates a structured evolution from a working static POC to a scalable, API-driven video generation capability. By following a phased MVP approach, the initiative balances speed, validation, and long-term extensibility.
This document serves as the foundational specification for ongoing development and stakeholder alignment.