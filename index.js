#!/usr/bin/env node

/**
 * Auto Resume Builder v1.0
 * Generate professional DOCX resumes from structured JSON data
 * 
 * Usage:
 *   node index.js < input.json         # Read from stdin
 *   node index.js --file resume.json   # Read from file
 *   node index.js --demo               # Generate demo resume
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageOrientation, PageSize, Tab, TabStopPosition, TabStopType,
  TableOfContents, ExternalHyperlink, HyperlinkType
} = require('docx');

// =============== STYLE CONSTANTS ===============

const COLORS = {
  primary: '1B3A5C',      // Dark navy blue
  accent: '2B7A78',       // Teal
  lightBg: 'F0F4F8',      // Light gray-blue
  darkText: '2D3436',     // Near black
  mediumText: '636E72',   // Gray
  lightText: 'FFFFFF',    // White
  divider: 'B2BEC3',      // Light gray
  sectionBg: 'E8F0FE',    // Light blue
};

// =============== RESUME BUILDER ===============

class ResumeBuilder {
  constructor(data) {
    this.data = data;
  }

  build() {
    const sections = [];

    // Header: Name + Title
    sections.push(this.createHeader());
    sections.push(this.createDivider());

    // Contact info
    if (this.data.contact) {
      sections.push(this.createContactSection());
      sections.push(this.createDivider());
    }

    // Professional Summary
    if (this.data.summary) {
      sections.push(this.createSectionTitle('Professional Summary'));
      sections.push(this.createSummarySection());
      sections.push(new Paragraph({ spacing: { after: 200 } }));
    }

    // Work Experience
    if (this.data.experience && this.data.experience.length > 0) {
      sections.push(this.createSectionTitle('Work Experience'));
      sections.push(...this.createExperienceSections());
    }

    // Education
    if (this.data.education && this.data.education.length > 0) {
      sections.push(this.createSectionTitle('Education'));
      sections.push(...this.createEducationSections());
    }

    // Skills
    if (this.data.skills && this.data.skills.length > 0) {
      sections.push(this.createSectionTitle('Skills'));
      sections.push(this.createSkillsSection());
    }

    // Certifications
    if (this.data.certifications && this.data.certifications.length > 0) {
      sections.push(this.createSectionTitle('Certifications'));
      sections.push(...this.createCertificationSections());
    }

    // Projects
    if (this.data.projects && this.data.projects.length > 0) {
      sections.push(this.createSectionTitle('Projects'));
      sections.push(...this.createProjectSections());
    }

    // Languages
    if (this.data.languages && this.data.languages.length > 0) {
      sections.push(this.createSectionTitle('Languages'));
      sections.push(this.createLanguagesSection());
    }

    return new Document({
      title: `${this.data.name || 'Resume'} - Resume`,
      description: `Professional resume for ${this.data.name || 'Candidate'}`,
      styles: {
        default: {
          document: {
            run: {
              font: 'Calibri',
              size: 22,
              color: COLORS.darkText,
            },
            paragraph: {
              spacing: { after: 80, line: 276 },
            },
          },
        },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 }, // Letter size
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
          },
        },
        children: sections,
      }],
    });
  }

  createHeader() {
    const name = this.data.name || 'Your Name';
    const title = this.data.title || '';

    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [
        new TextRun({
          text: name,
          bold: true,
          size: 52,
          color: COLORS.primary,
          font: 'Calibri',
        }),
      ],
    });
  }

  createDivider() {
    return new Paragraph({
      spacing: { before: 100, after: 100 },
      border: {
        bottom: { color: COLORS.primary, size: 6, style: BorderStyle.SINGLE, space: 1 },
      },
    });
  }

  createContactSection() {
    const c = this.data.contact;
    if (!c) return new Paragraph({});

    const parts = [];
    if (c.phone) parts.push(c.phone);
    if (c.email) parts.push(c.email);
    if (c.location) parts.push(c.location);

    const line = parts.join('  |  ');

    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [
        new TextRun({
          text: line,
          size: 20,
          color: COLORS.mediumText,
          font: 'Calibri',
        }),
      ],
    });
  }

  createSectionTitle(text) {
    return new Paragraph({
      spacing: { before: 240, after: 120 },
      children: [
        new TextRun({
          text: text.toUpperCase(),
          bold: true,
          size: 24,
          color: COLORS.primary,
          font: 'Calibri',
        }),
      ],
      border: {
        bottom: { color: COLORS.accent, size: 4, style: BorderStyle.SINGLE, space: 1 },
      },
    });
  }

  createSummarySection() {
    return new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: this.data.summary,
          size: 22,
          color: COLORS.darkText,
          font: 'Calibri',
        }),
      ],
    });
  }

  createExperienceSections() {
    const sections = [];
    for (const exp of this.data.experience) {
      // Job header: company + dates
      const dateRange = `${exp.startDate || ''} - ${exp.endDate || 'Present'}`;
      sections.push(
        new Paragraph({
          spacing: { before: 120, after: 40 },
          children: [
            new TextRun({
              text: exp.company || '',
              bold: true,
              size: 24,
              color: COLORS.darkText,
              font: 'Calibri',
            }),
            new TextRun({
              text: `\t${dateRange}`,
              size: 20,
              color: COLORS.mediumText,
              font: 'Calibri',
            }),
          ],
        })
      );

      // Position
      if (exp.position) {
        sections.push(
          new Paragraph({
            spacing: { before: 20, after: 60 },
            children: [
              new TextRun({
                text: exp.position,
                size: 22,
                color: COLORS.accent,
                font: 'Calibri',
                italics: true,
              }),
            ],
          })
        );
      }

      // Description bullets
      if (exp.description && exp.description.length > 0) {
        for (const desc of exp.description) {
          sections.push(
            new Paragraph({
              spacing: { after: 40 },
              indent: { left: 400 },
              children: [
                new TextRun({
                  text: '•  ',
                  size: 22,
                  color: COLORS.accent,
                  font: 'Calibri',
                }),
                new TextRun({
                  text: desc,
                  size: 22,
                  color: COLORS.darkText,
                  font: 'Calibri',
                }),
              ],
            })
          );
        }
      }
    }
    return sections;
  }

  createEducationSections() {
    const sections = [];
    for (const edu of this.data.education) {
      sections.push(
        new Paragraph({
          spacing: { before: 100, after: 40 },
          children: [
            new TextRun({
              text: edu.institution || '',
              bold: true,
              size: 23,
              color: COLORS.darkText,
              font: 'Calibri',
            }),
            new TextRun({
              text: `\t${edu.graduationYear || ''}`,
              size: 20,
              color: COLORS.mediumText,
              font: 'Calibri',
            }),
          ],
        })
      );

      if (edu.degree) {
        sections.push(
          new Paragraph({
            spacing: { after: 40 },
            indent: { left: 200 },
            children: [
              new TextRun({
                text: edu.degree,
                size: 22,
                color: COLORS.accent,
                font: 'Calibri',
              }),
              ...(edu.field ? [
                new TextRun({
                  text: ` in ${edu.field}`,
                  size: 22,
                  color: COLORS.darkText,
                  font: 'Calibri',
                }),
              ] : []),
            ],
          })
        );
      }

      if (edu.gpa) {
        sections.push(
          new Paragraph({
            spacing: { after: 60 },
            indent: { left: 200 },
            children: [
              new TextRun({
                text: `GPA: ${edu.gpa}`,
                size: 20,
                color: COLORS.mediumText,
                font: 'Calibri',
              }),
            ],
          })
        );
      }
    }
    return sections;
  }

  createSkillsSection() {
    const items = this.data.skills.map(s => {
      if (typeof s === 'string') return s;
      if (typeof s === 'object' && s.name) return `${s.name}${s.level ? ` (${s.level})` : ''}`;
      return String(s);
    });

    // Split into groups of 3 for columns
    const groups = [];
    for (let i = 0; i < items.length; i += 3) {
      groups.push(items.slice(i, i + 3));
    }

    const paragraphs = [];
    for (const group of groups) {
      paragraphs.push(
        new Paragraph({
          spacing: { after: 40 },
          indent: { left: 200 },
          children: [
            new TextRun({
              text: '•  ',
              size: 22,
              color: COLORS.accent,
              font: 'Calibri',
            }),
            new TextRun({
              text: group.join('    •  '),
              size: 22,
              color: COLORS.darkText,
              font: 'Calibri',
            }),
          ],
        })
      );
    }
    return paragraphs.length > 0 ? paragraphs[0] : new Paragraph({});
  }

  createCertificationSections() {
    return this.data.certifications.map(cert =>
      new Paragraph({
        spacing: { after: 40 },
        indent: { left: 200 },
        children: [
          new TextRun({
            text: '•  ',
            size: 22,
            color: COLORS.accent,
            font: 'Calibri',
          }),
          new TextRun({
            text: typeof cert === 'string' ? cert : cert.name || '',
            size: 22,
            color: COLORS.darkText,
            font: 'Calibri',
          }),
          ...(typeof cert === 'object' && cert.issuer ? [
            new TextRun({
              text: ` - ${cert.issuer}`,
              size: 20,
              color: COLORS.mediumText,
              font: 'Calibri',
            }),
          ] : []),
        ],
      })
    );
  }

  createProjectSections() {
    return this.data.projects.flatMap(proj => {
      const items = [
        new Paragraph({
          spacing: { before: 80, after: 20 },
          children: [
            new TextRun({
              text: proj.name || '',
              bold: true,
              size: 23,
              color: COLORS.darkText,
              font: 'Calibri',
            }),
            ...(proj.url ? [
              new TextRun({
                text: `  (${proj.url})`,
                size: 18,
                color: COLORS.accent,
                font: 'Calibri',
              }),
            ] : []),
          ],
        }),
      ];

      if (proj.description) {
        items.push(
          new Paragraph({
            spacing: { after: 40 },
            indent: { left: 200 },
            children: [
              new TextRun({
                text: proj.description,
                size: 22,
                color: COLORS.darkText,
                font: 'Calibri',
              }),
            ],
          })
        );
      }

      return items;
    });
  }

  createLanguagesSection() {
    return new Paragraph({
      spacing: { after: 40 },
      indent: { left: 200 },
      children: [
        new TextRun({
          text: this.data.languages.map(l =>
            typeof l === 'string' ? l : `${l.name}${l.level ? ` (${l.level})` : ''}`
          ).join('    •  '),
          size: 22,
          color: COLORS.darkText,
          font: 'Calibri',
        }),
      ],
    });
  }
}

// =============== CLI INTERFACE ===============

function printBanner() {
  console.log(`
  ╔══════════════════════════════════════════╗
  ║     Auto Resume Builder v1.0             ║
  ║     Generate professional DOCX resumes   ║
  ╚══════════════════════════════════════════╝
  `);
}

function showUsage() {
  console.log(`
  Usage:
    node index.js < data.json          Generate from stdin
    node index.js --file data.json     Generate from file
    node index.js --input data.json    Generate from file
    node index.js --demo               Generate demo resume
    node index.js --help               Show this help

  JSON Input Format:
    {
      "name": "Zhang Wei",
      "title": "Senior Software Engineer",
      "contact": { "phone": "+86 138-0000-0000", "email": "zhangwei@email.com", "location": "Beijing, China" },
      "summary": "...",
      "experience": [
        { "company": "ABC Tech", "position": "Senior Developer",
          "startDate": "2020-01", "endDate": "Present",
          "description": ["Led team of 5 engineers...", "Developed microservices..."] }
      ],
      "education": [...],
      "skills": ["JavaScript", "Python", "React", ...],
      "certifications": [...],
      "projects": [...],
      "languages": [...]
    }
  `);
}

async function main() {
  const args = process.argv.slice(2);
  let inputData = null;

  if (args.includes('--help') || args.includes('-h')) {
    printBanner();
    showUsage();
    process.exit(0);
  }

  if (args.includes('--demo')) {
    inputData = {
      name: 'Zhang Wei',
      title: 'Senior Software Engineer',
      contact: { phone: '+86 138-0000-0000', email: 'zhangwei@email.com', location: 'Beijing, China' },
      summary: 'Experienced software engineer with 8+ years of expertise in full-stack development, cloud architecture, and team leadership. Passionate about building scalable systems and mentoring junior developers.',
      experience: [
        {
          company: 'Tech Innovations Inc.',
          position: 'Senior Software Engineer',
          startDate: '2020-03',
          endDate: 'Present',
          description: [
            'Led a team of 5 engineers to architect and implement microservices-based platform, reducing deployment time by 60%',
            'Designed and developed RESTful APIs handling 10M+ daily requests with 99.9% uptime',
            'Implemented CI/CD pipelines using GitHub Actions and Docker, achieving 4x faster release cycles',
            'Mentored 3 junior developers through code reviews and pair programming sessions',
          ],
        },
        {
          company: 'DataFlow Solutions',
          position: 'Full Stack Developer',
          startDate: '2017-06',
          endDate: '2020-02',
          description: [
            'Built real-time data processing dashboard using React and Node.js, serving 500+ enterprise clients',
            'Optimized database queries reducing average response time from 2s to 200ms',
            'Migrated legacy monolith to cloud-native architecture on AWS, reducing infrastructure costs by 40%',
          ],
        },
        {
          company: 'WebCraft Studio',
          position: 'Junior Developer',
          startDate: '2015-09',
          endDate: '2017-05',
          description: [
            'Developed responsive web applications using React, Vue.js, and Express',
            'Collaborated with design team to implement pixel-perfect UI components',
          ],
        },
      ],
      education: [
        { institution: 'Beijing University of Technology', degree: 'Bachelor of Science', field: 'Computer Science', graduationYear: '2015', gpa: '3.7/4.0' },
      ],
      skills: ['JavaScript/TypeScript', 'Python', 'React/Next.js', 'Node.js/Express', 'Docker/Kubernetes', 'AWS/GCP', 'PostgreSQL/MongoDB', 'GraphQL', 'Redis', 'Git/GitHub Actions'],
      certifications: [
        { name: 'AWS Solutions Architect Professional', issuer: 'Amazon Web Services' },
        { name: 'Google Cloud Professional Data Engineer', issuer: 'Google Cloud' },
      ],
      projects: [
        { name: 'Open Source Monitoring Tool', description: 'Built and maintained an open-source server monitoring dashboard with 2K+ GitHub stars', url: 'github.com/example/monitor' },
      ],
      languages: ['Chinese (Native)', 'English (Fluent - IELTS 7.5)'],
    };
  } else {
    // Read from file or stdin
    const fileFlag = args.indexOf('--file') > -1 ? '--file' : (args.indexOf('--input') > -1 ? '--input' : null);
    if (fileFlag) {
      const fileIdx = args.indexOf(fileFlag) + 1;
      if (fileIdx >= args.length) {
        console.error('Error: Missing file path');
        process.exit(1);
      }
      const filePath = args[fileIdx];
      try {
        inputData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      } catch (e) {
        console.error(`Error reading file: ${e.message}`);
        process.exit(1);
      }
    } else {
      // Read from stdin
      try {
        const stdin = fs.readFileSync('/dev/stdin', 'utf8');
        if (!stdin || stdin.trim().length === 0) {
          printBanner();
          showUsage();
          process.exit(0);
        }
        inputData = JSON.parse(stdin);
      } catch (e) {
        console.error('Error: No input data. Use --demo for a sample, or pipe JSON data.');
        showUsage();
        process.exit(1);
      }
    }
  }

  printBanner();
  console.log(`Generating resume for: ${inputData.name || 'Candidate'}`);

  const builder = new ResumeBuilder(inputData);
  const doc = builder.build();

  const outputName = `${(inputData.name || 'Resume').replace(/\s+/g, '_')}_Resume.docx`;
  const outputPath = path.resolve(outputName);

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);

  console.log(`\n✅ Resume generated: ${outputPath}`);

  const stats = fs.statSync(outputPath);
  console.log(`   Size: ${(stats.size / 1024).toFixed(1)} KB`);
}

main().catch(console.error);
