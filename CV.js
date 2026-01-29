const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, LevelFormat, TabStopType, TabStopPosition, ExternalHyperlink } = require('docx');
const fs = require('fs');
const docxPdf = require('docx-pdf');

const doc = new Document({
  styles: {
    default: { 
      document: { 
        run: { font: "Arial", size: 22 } // 11pt default for better space usage
      } 
    },
    paragraphStyles: [
      { 
        id: "Heading1", 
        name: "Heading 1", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" }, // 14pt
        paragraph: { 
          spacing: { before: 240, after: 120 }, 
          outlineLevel: 0 
        } 
      },
      { 
        id: "Heading2", 
        name: "Heading 2", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 26, bold: true, font: "Arial" }, // 13pt
        paragraph: { 
          spacing: { before: 180, after: 100 }, 
          outlineLevel: 1 
        } 
      },
      { 
        id: "Heading3", 
        name: "Heading 3", 
        basedOn: "Normal", 
        next: "Normal", 
        quickFormat: true,
        run: { size: 24, bold: true, font: "Arial" }, // 12pt
        paragraph: { 
          spacing: { before: 120, after: 60 }, 
          outlineLevel: 2 
        } 
      },
    ]
  },
  numbering: {
    config: [
      { 
        reference: "bullets",
        levels: [
          { 
            level: 0, 
            format: LevelFormat.BULLET, 
            text: "•", 
            alignment: AlignmentType.LEFT,
            style: { 
              paragraph: { 
                indent: { left: 720, hanging: 360 },
                spacing: { after: 60 }
              } 
            } 
          }
        ] 
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: {
          width: 12240,   // US Letter
          height: 15840
        },
        margin: { top: 2160, right: 2160, bottom: 2160, left: 2160 } // 1.5 inch margins
      }
    },
    children: [
      // NAME AND CONTACT (Header)
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "YASSER CHAOUKI",
            bold: true,
            size: 55, // 16pt
            font: "Arial"
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: "+212 707512723 | ",
            size: 20 // 10pt
          }),
          new ExternalHyperlink({
            children: [
              new TextRun({
                text: "chaouki0yasser5@gmail.com",
                size: 20,
                style: "Hyperlink"
              })
            ],
            link: "mailto:chaouki0yasser5@gmail.com"
          }),
          new TextRun({
            text: " | ",
            size: 20
          }),
          new ExternalHyperlink({
            children: [
              new TextRun({
                text: "https://www.linkedin.com/in/yasser-chaouki",
                size: 20,
                style: "Hyperlink"
              })
            ],
            link: "https://www.linkedin.com/in/yasser-chaouki"
          }),
          new TextRun({
            text: " | Tangier, Morocco",
            size: 20
          })
        ]
      }),

      // STRATEGIC SUMMARY
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("STRATEGIC SUMMARY")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: "I am a highly motivated first-year Cybersecurity engineering student with a robust foundation in mathematics, algorithm design, and programming (Python, Java). I bring specialized expertise in cryptography and information security, demonstrated through independent research on RSA cryptosystems, lattice-based attacks, and cryptanalytic vulnerabilities. My analytical approach combines theoretical rigor—applying number theory and algebraic proofs to security problems—with practical implementation skills. I'm passionate about cyber defense and interested in exploring machine learning and automation applications within cybersecurity. I'm continuously expanding my knowledge through hands-on learning courses. I'm eager to contribute my theoretical knowledge, research experience, and technical skills to a challenging cybersecurity internship where I can apply my foundation in algorithm design and information security while developing capabilities in machine learning and automation. My goal is to leverage my analytical mindset and programming skills to help organizations strengthen their security posture against evolving cyber threats."
          })
        ]
      }),

      // EDUCATION
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("EDUCATION")]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Ecole Nationale Supérieure de l'Intelligence Artificielle et Science de Données (ENSIASD)",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Engineering degree in cybersecurity | Sept 2025 – Present | Taroudant, Morocco"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Specialization: Adversarial machine learning, secure AI architectures, network security"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Relevant coursework: Cryptography, algorithmic complexity, network security, probability & statistics, python, machine learning, reseau informatique, architecture des ordinateurs"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Project achievement: Engineered predictive ML model using Python and Scikit-learn (Dec 2025), achieving 92% accuracy in classifying network intrusions based on real-world datasets (January 2026), started a RAG system for mathematical problem solving using LlamaIndex and Pinecone (January 2026), started a 'learning by doing' github repo from microsoft titeled ML-For-Beginners (January 2026)"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "CPGE Moulay Al Hassan",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Mathematics & Physics (MPSI/MP) | Sept 2023 – July 2025 | Morocco"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Intensive two-year program: Advanced linear algebra, multivariable calculus, discrete mathematics"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Performance: Consistently ranked top student in Mathematics and Python programming"
          })
        ]
      }),

      // CERTIFICATIONS & AWARDS
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240 },
        children: [new TextRun("CERTIFICATIONS & AWARDS")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "ISC2 certified in cybersecurity (CC): Candidate (In progress)"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Math&Maroc competition (MMC) 2024: National finalist – Recognized for problem-solving and mathematical reasoning skills"
          })
        ]
      }),

      // TECHNICAL SKILLS
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240 },
        children: [new TextRun("TECHNICAL SKILLS")]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Security & cryptography: ",
            bold: true
          }),
          new TextRun({
            text: "RSA cryptanalysis, Integer factorization, network security, linux terminal, VirtualBox/VMware, TCP/IP, DNS, VPNs"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "AI/Machine Learning (In progress): ",
            bold: true
          }),
          new TextRun({
            text: "LLMs (Llama-3, GPT), RAG systems (LangChain, LlamaIndex, Pinecone, ChromaDB), neural networks, computer vision, PyTorch/TensorFlow, Scikit-learn, NumPy, Pandas, Supervised/Unsupervised Learning, feature engineering, model validation"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Programming & Development: ",
            bold: true
          }),
          new TextRun({
            text: "Python (Expert), Java (Object-Oriented Design), SQL, C/C++, JavaScript(In progress), FastAPI(In progress), Git, Bash, n8n Automation(In progress)"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Mathematical Foundations: ",
            bold: true
          }),
          new TextRun({
            text: "Linear algebra, multivariable calculus, probability theory, discrete mathematics, graph theory, algorithm design, complexity analysis (Big O), stochastic optimization"
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: "Tools & Environments: ",
            bold: true
          }),
          new TextRun({
            text: "Linux/Unix Shell, VS Code, Jupyter Notebooks, Figma, LaTeX-OCR"
          })
        ]
      }),

      // CYBERSECURITY RESEARCH
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("CYBERSECURITY RESEARCH")]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "Cryptographic research analyst (TIPE)",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "CPGE | Jan 2023 – Jun 2024"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "RSA security audit: Conducted comprehensive theoretical audit of RSA public-key cryptosystem, analyzing computational complexity of integer factorization through rigorous mathematical investigation"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Attack simulation: Designed and simulated lattice-based side-channel attacks using Graham-Schmidt orthogonalization for polynomial reduction, demonstrating RSA key generation vulnerabilities"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Mathematical framework: Formulated algebraic proofs bridging abstract number theory with practical PKI security standards, establishing theoretical foundation for cryptanalytic vulnerabilities"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Publication: Authored comprehensive research paper on cryptanalytic attack frameworks, published on LinkedIn for technical community engagement"
          })
        ]
      }),

      // AI & SOFTWARE ENGINEERING PROJECTS
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240 },
        children: [new TextRun("AI & SOFTWARE ENGINEERING PROJECTS")]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({
            text: "AI research engineer – Symbolic reasoning & RAG systems (In progress)",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Independent R&D | Jan 2025 – Present"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "RAG architecture: Engineering advanced retrieval-augmented generation (RAG) system optimized for symbolic mathematics and pedagogical reasoning, benchmarked against CPGE-level problem sets"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Technical stack: Orchestrated workflows using LangChain and n8n, implementing Pinecone/ChromaDB for mathematical theorem indexing. Integrated LaTeX-OCR for PDF equation parsing"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "LLM fine-tuning: Fine-tuned Llama-3-70B with custom Chain-of-Thought (CoT) system prompts, achieving 15% improvement in reasoning accuracy on GSM8K benchmark via multi-shot prompting"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Performance optimization: Optimized retrieval latency to <200ms using HNSW indexing and implemented self-correction agent reducing LaTeX rendering errors by 40%"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "Lead software architect & project manager",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Enactus ENSIASD | Taroudant, Morocco | Sept 2025 – Present"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Strategic innovation: Leading end-to-end development of hyperlocal delivery infrastructure, bridging logistical gaps between Taroudant vendors and ENSIASD student community"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Backend architecture: Architecting robust backend system for real-time order processing and dispatching logic, ensuring high availability for 100+ student users"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Agile leadership: Managing full SDLC from requirement gathering to MVP deployment, leading cross-functional team using \"pedagogy of doing\" principles"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "AI automation engineer",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Independent technical projects | 2024 – 2025"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Academic planner agent: Engineered custom AI agent using Google Gemini API with logic-based validation loops and API key security protocols"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Math tutoring agent: Developed specialized tutoring agent for CPGE students, optimizing prompt engineering for logical verification and step-by-step problem solving"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Travel planner agent: Built autonomous travel planning agent with API orchestration, data minimization protocols, and secure key management"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "Computer Vision & Neural Network Development",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Independent R&D | 2024"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Neural network from scratch: Built and trained neural network using Python to classify alphanumeric datasets, applying Linear Algebra and Calculus principles for weight optimization and backpropagation"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Model validation: Implemented performance metrics and validation techniques to ensure model accuracy and generalization"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 180, after: 60 },
        children: [
          new TextRun({
            text: "Full Stack Developer",
            bold: true
          })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: "Freelance & Portfolio Projects | 2024"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Java enterprise application: Developed comprehensive stock management and invoicing application using Java/Swing, implementing strict object-oriented design patterns for modularity and maintainability"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Rapid prototyping: Delivered full-stack web solutions for local businesses, designing high-fidelity Figma prototypes and translating them into responsive frontend code with functional backend integration"
          })
        ]
      }),

      // PROFESSIONAL ATTRIBUTES
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240 },
        children: [new TextRun("PROFESSIONAL ATTRIBUTES")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Analytical thinking: Proven ability to deconstruct complex cybersecurity problems using first-principles reasoning and mathematical rigor"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Autonomous learning: Self-directed mastery of ML stack, cryptographic systems, and AI frameworks through project-based learning"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Leadership & collaboration: Cross-functional team leadership, peer-to-peer learning, and collective intelligence in Agile environments"
          })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [
          new TextRun({
            text: "Resilience: High-pressure performance in competitive mathematics and intensive CPGE curriculum, demonstrating adaptability and commitment to excellence"
          })
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const docxPath = "Yasser_Chaouki_CV_ATS_Optimized.docx";
  const pdfPath = "Yasser_Chaouki_CV_ATS_Optimized.pdf";
  fs.writeFileSync(docxPath, buffer);
  docxPdf(docxPath, pdfPath, (err) => {
    if (err) {
      console.error("Error converting to PDF:", err);
    } else {
      console.log("PDF created successfully!");
    }
  });
});