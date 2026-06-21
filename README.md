# Auto Resume Builder

AI-powered resume builder — Generate professional DOCX resumes from simple JSON data.

## Quick Start

```bash
# Generate demo resume
node index.js --demo

# Generate from JSON file
node index.js --file resume.json
```

## Install

```bash
npm install
```

## Input Format

Create a JSON file with your information:

```json
{
  "name": "Zhang Wei",
  "title": "Senior Software Engineer",
  "contact": {
    "phone": "+86 138-0000-0000",
    "email": "zhangwei@email.com",
    "location": "Beijing, China"
  },
  "summary": "Experienced software engineer with 8+ years...",
  "experience": [
    {
      "company": "Tech Co.",
      "position": "Senior Developer",
      "startDate": "2020-01",
      "endDate": "Present",
      "description": [
        "Led team of 5 engineers...",
        "Built microservices platform..."
      ]
    }
  ],
  "education": [
    {
      "institution": "University",
      "degree": "Bachelor",
      "field": "Computer Science",
      "graduationYear": "2015"
    }
  ],
  "skills": ["JavaScript", "Python", "React"],
  "certifications": [
    { "name": "AWS Certified", "issuer": "Amazon" }
  ],
  "languages": ["English (Fluent)", "Chinese (Native)"]
}
```

## Features

- Professional layout with clean typography
- ATS-friendly output format
- Customizable sections
- 100% local, no external services

---

## ☕ 赞赏支持

如果这个工具帮到了你，欢迎请我喝杯咖啡 ❤️

| 微信扫一扫赞赏 | 
|:--------------:|
| ![微信赞赏码](donate.jpg) |

所有赞赏都会用于持续改进和开发更多免费工具 🙏

## License

MIT
