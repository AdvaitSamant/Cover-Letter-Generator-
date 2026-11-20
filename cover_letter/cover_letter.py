import asyncio
import io
import re
from functools import lru_cache
from typing import TYPE_CHECKING

import reflex as rx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

if TYPE_CHECKING:
    from transformers import TextGenerationPipeline

MODEL = "Qwen/Qwen2.5-1.5B-Instruct"

TEMPLATES = {
    "professional": {
        "name": "Professional Classic",
        "tone": "formal, traditional business letter format with a measured, professional tone",
        "font": "Calibri",
        "accent": "#1e1b4b",
    },
    "modern": {
        "name": "Modern Clean",
        "tone": "clean, contemporary professional tone with active voice and concise language",
        "font": "Calibri Light",
        "accent": "#0d9488",
    },
    "creative": {
        "name": "Creative Bold",
        "tone": "engaging and distinctive professional tone that shows personality while remaining appropriate",
        "font": "Arial",
        "accent": "#dc2626",
    },
    "executive": {
        "name": "Executive Formal",
        "tone": "highly polished, executive-level tone suitable for senior leadership roles, emphasizing strategic impact and vision",
        "font": "Georgia",
        "accent": "#374151",
    },
}


def build_prompt(resume: str, jd: str, company: str, tone: str) -> str:
    return f"""You are an expert career coach and professional cover letter writer.

Using ONLY the real details from the resume below, write a compelling, tailored cover letter for the job description provided. The letter must be addressed to the hiring team at **{company}**.

### STYLE
Write in a {tone}.

### RULES
1. Use ONLY facts, skills, experiences, and achievements that appear in the resume. Do NOT invent anything.
2. Map the candidate's strengths directly to the job requirements.
3. The letter should be approximately 250 words.
4. Use standard cover letter structure: opening hook, relevant experience/skills, why this company, closing call to action.
5. Do NOT include any placeholders like "[Your Name]" — extract names and details from the resume.
6. Output ONLY the cover letter text — no headings, no labels, no markdown formatting.

---

### RESUME
{resume}

---

### JOB DESCRIPTION
{jd}

---

Write the cover letter now."""


@lru_cache(maxsize=1)
def _load_text_generator() -> "TextGenerationPipeline":
    import torch
    from transformers import AutoModelForCausalLM, AutoTokenizer, pipeline

    if not torch.cuda.is_available():
        raise RuntimeError(
            "CUDA is not available. This application is configured to run on GPU only."
        )

    tokenizer = AutoTokenizer.from_pretrained(MODEL)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL,
        dtype=torch.float16,
        low_cpu_mem_usage=True,
        device_map="cuda",
    )
    return pipeline(
        "text-generation",
        model=model,
        tokenizer=tokenizer,
    )


def generate_cover_letter(resume: str, jd: str, company: str, tone: str) -> str:
    generator = _load_text_generator()
    prompt = build_prompt(resume, jd, company, tone)
    if hasattr(generator.tokenizer, "chat_template") and generator.tokenizer.chat_template:
        messages = [{"role": "user", "content": prompt}]
        prompt = generator.tokenizer.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )

    generated = generator(
        prompt,
        do_sample=True,
        temperature=0.7,
        max_new_tokens=512,
        return_full_text=False,
    )
    content = generated[0].get("generated_text")
    if isinstance(content, str) and content.strip():
        return content.strip()
    raise RuntimeError("The model returned no final cover letter text.")


def add_colored_line(doc: Document, color_hex: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(2)
    properties = paragraph._element.get_or_add_pPr()
    borders = properties.makeelement(qn("w:pBdr"), {})
    bottom = borders.makeelement(
        qn("w:bottom"),
        {
            qn("w:val"): "single",
            qn("w:sz"): "6",
            qn("w:space"): "1",
            qn("w:color"): color_hex.lstrip("#"),
        },
    )
    borders.append(bottom)
    properties.append(borders)


def create_docx(letter_text: str, company: str, template_key: str = "professional") -> io.BytesIO:
    template = TEMPLATES[template_key]
    accent_hex = template["accent"]
    font_name = template["font"]
    r, g, b = (
        int(accent_hex[1:3], 16),
        int(accent_hex[3:5], 16),
        int(accent_hex[5:7], 16),
    )

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title = doc.add_paragraph()
    title.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
        if template_key in ("professional", "executive")
        else WD_ALIGN_PARAGRAPH.LEFT
    )
    run = title.add_run(f"Cover Letter — {company}")
    run.bold = True
    run.font.size = Pt(16) if template_key != "executive" else Pt(18)
    run.font.color.rgb = RGBColor(r, g, b)
    run.font.name = font_name

    add_colored_line(doc, accent_hex)

    paragraphs = [p.strip() for p in re.split(r"\n{2,}", letter_text) if p.strip()]
    for para_text in paragraphs:
        paragraph = doc.add_paragraph()
        paragraph.alignment = (
            WD_ALIGN_PARAGRAPH.JUSTIFY
            if template_key == "executive"
            else WD_ALIGN_PARAGRAPH.LEFT
        )
        paragraph.paragraph_format.space_after = Pt(10)
        paragraph.paragraph_format.line_spacing = 1.5 if template_key == "executive" else 1.4
        run = paragraph.add_run(para_text)
        run.font.size = Pt(11)
        run.font.name = font_name
        run.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


class State(rx.State):
    company: str = ""
    resume: str = ""
    target_role: str = ""
    cover_letter: str = ""
    company_used: str = ""
    is_generating: bool = False
    generation_error: str = ""
    company_error: str = ""
    resume_error: str = ""
    role_error: str = ""
    selected_template: str = "professional"

    def set_template(self, template_key: str):
        self.selected_template = template_key

    def set_company_value(self, value: str):
        self.company = value

    def set_resume_value(self, value: str):
        self.resume = value

    def set_target_role_value(self, value: str):
        self.target_role = value

    async def generate(self):
        self.company_error = ""
        self.resume_error = ""
        self.role_error = ""
        self.generation_error = ""

        if not self.company.strip():
            self.company_error = "Please enter a company name."
        if not self.resume.strip():
            self.resume_error = "Please paste your resume."
        if not self.target_role.strip():
            self.role_error = "Please paste the job description."
        if self.company_error or self.resume_error or self.role_error:
            return

        self.is_generating = True
        yield

        try:
            letter = await asyncio.to_thread(
                generate_cover_letter,
                self.resume,
                self.target_role,
                self.company,
                TEMPLATES[self.selected_template]["tone"],
            )
            self.cover_letter = letter
            self.company_used = self.company
        except Exception as exc:
            self.generation_error = f"Could not load or run the local model: {exc}"
            self.cover_letter = ""
        finally:
            self.is_generating = False

    def download_docx(self):
        if not self.cover_letter:
            return
        safe_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", self.company_used).strip().replace(" ", "_")
        document = create_docx(self.cover_letter, self.company_used, self.selected_template)
        return rx.download(
            data=document.getvalue(),
            filename=f"Cover_Letter_{safe_name}.docx",
            mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    def copy_letter(self):
        return rx.set_clipboard(self.cover_letter)


def nav() -> rx.Component:
    return rx.el.nav(
        rx.el.div(
            rx.el.div(
                rx.el.span("✉", class_name="brand-mark"),
                rx.el.span("Cover Letter Studio", class_name="brand-name"),
                class_name="brand",
            ),
            rx.el.div(
                rx.el.a("Home", href="#home"),
                rx.el.a("How it works", href="#how-it-works"),
                rx.el.a("About", href="#about"),
                class_name="nav-links",
            ),
            class_name="nav-inner",
        ),
        class_name="studio-nav",
    )


def hero() -> rx.Component:
    return rx.el.section(
        rx.el.div("Your next opportunity starts here", class_name="hero-eyebrow"),
        rx.el.h1(
            rx.el.span("Your experience.", class_name="line"),
            rx.el.span("Your story.", class_name="line"),
            rx.el.span("Your next opportunity.", class_name="accent"),
        ),
        rx.el.p(
            "Transform your resume into a compelling cover letter, tailored to the role you're applying for. "
            "No generic templates. Just your experience, presented with purpose.",
            class_name="hero-copy",
        ),
        id="home",
        class_name="hero",
    )


def card_header(number: str, eyebrow: str, subtitle: str) -> rx.Component:
    return rx.el.div(
        rx.el.div(rx.el.strong(number), rx.el.span(eyebrow), class_name="card-kicker"),
        rx.el.p(subtitle, class_name="card-subtitle"),
        class_name="card-header",
    )


def ghost_letter() -> rx.Component:
    return rx.el.div(
        *[rx.el.div(class_name="ghost-line short" if i in (1, 4) else "ghost-line") for i in range(6)],
        class_name="ghost-letter",
        aria_hidden="true",
    )


def template_button(key: str, icon: str, name: str, desc: str) -> rx.Component:
    is_selected = State.selected_template == key
    return rx.el.button(
        rx.el.div(
            rx.el.span(icon, class_name="template-icon"),
            rx.el.div(
                rx.el.strong(name, class_name="template-name"),
                rx.el.span(desc, class_name="template-desc"),
                class_name="template-text",
            ),
            class_name="template-inner",
        ),
        type="button",
        on_click=State.set_template(key),
        class_name=rx.cond(is_selected, "template-card active", "template-card"),
    )


def template_selector() -> rx.Component:
    return rx.el.div(
        rx.el.label("Tone & Template Style"),
        rx.el.div(
            template_button("professional", "💼", "Professional Classic", "Formal business tone · Calibri"),
            template_button("modern", "⚡", "Modern Clean", "Active & concise · Calibri Light"),
            template_button("creative", "🎨", "Creative Bold", "Engaging personality · Arial"),
            template_button("executive", "🏛️", "Executive Formal", "Strategic leadership · Georgia"),
            class_name="template-grid",
        ),
        class_name="template-section",
    )


def left_card() -> rx.Component:
    return rx.el.div(
        card_header("01", "Your details", "Give us the context. We’ll handle the writing."),
        rx.el.div(
            rx.el.label("Company"),
            rx.input(
                placeholder="e.g. Google, Microsoft, NVIDIA",
                value=State.company,
                on_change=State.set_company_value,
                class_name="field",
            ),
            rx.cond(State.company_error != "", rx.el.div(State.company_error, class_name="field-error")),
            rx.el.label("Your resume"),
            rx.text_area(
                placeholder="Paste your resume here. Include your experience, education, projects, and skills...",
                value=State.resume,
                on_change=State.set_resume_value,
                height="220px",
                class_name="field text-area",
            ),
            rx.cond(State.resume_error != "", rx.el.div(State.resume_error, class_name="field-error")),
            rx.el.label("Target role"),
            rx.text_area(
                placeholder="Paste the job description for the position you're applying to...",
                value=State.target_role,
                on_change=State.set_target_role_value,
                height="170px",
                class_name="field text-area",
            ),
            rx.cond(State.role_error != "", rx.el.div(State.role_error, class_name="field-error")),
            template_selector(),
            rx.button(
                rx.cond(State.is_generating, "Crafting your story…", "✦  Generate my cover letter  →"),
                on_click=State.generate,
                disabled=State.is_generating,
                class_name="generate-button",
                width="100%",
            ),
            rx.el.p(
                "Powered by Qwen2.5 1.5B · GPU Accelerated",
                class_name="powered-note",
            ),
            class_name="left-fields",
        ),
        class_name="studio-card left-card",
    )


def right_card() -> rx.Component:
    return rx.el.div(
        card_header("02", "Your cover letter", "A first impression worth remembering."),
        rx.cond(
            State.is_generating,
            rx.el.div(
                *[rx.el.div(class_name="skeleton short" if i == 2 else "skeleton") for i in range(5)],
                rx.el.div("Finding the right words…", class_name="loading-caption"),
                class_name="right-state loading-state",
            ),
            rx.cond(
                State.generation_error != "",
                rx.el.div(
                    ghost_letter(),
                    rx.el.h3("We couldn’t finish the letter.", class_name="empty-heading"),
                    rx.el.p(State.generation_error, class_name="empty-copy"),
                    rx.button("Retry", on_click=State.generate, class_name="retry-button"),
                    class_name="right-state",
                ),
                rx.cond(
                    State.cover_letter != "",
                    rx.el.div(
                        rx.el.div("Generated cover letter", class_name="preview-kicker"),
                        rx.el.div(
                            rx.foreach(
                                State.cover_letter.split("\n\n"),
                                lambda paragraph: rx.el.p(paragraph),
                            ),
                            class_name="doc-page",
                        ),
                        rx.el.div(
                            rx.button("Download .DOCX", on_click=State.download_docx, class_name="download-button"),
                            rx.button("Copy text", on_click=State.copy_letter, class_name="copy-button"),
                            rx.button("Regenerate", on_click=State.generate, class_name="regenerate-button"),
                            class_name="result-actions",
                        ),
                        class_name="result-state",
                    ),
                    rx.el.div(
                        ghost_letter(),
                        rx.el.h3("Your next great introduction awaits.", class_name="empty-heading"),
                        rx.el.p(
                            "Add your details and generate a personalized cover letter. Your finished letter will appear here.",
                            class_name="empty-copy",
                        ),
                        class_name="right-state",
                    ),
                ),
            ),
        ),
        class_name="studio-card right-card",
    )


def workspace() -> rx.Component:
    return rx.el.div(left_card(), right_card(), class_name="workspace")


def sections() -> rx.Component:
    return rx.el.div(
        rx.el.section(
            rx.el.h2("How it works"),
            rx.el.div(
                rx.el.article(
                    rx.el.div("01 / CONTEXT", class_name="step-number"),
                    rx.el.h3("Add your experience"),
                    rx.el.p("Paste your resume and the role you’re pursuing."),
                    class_name="step",
                ),
                rx.el.article(
                    rx.el.div("02 / ALIGNMENT", class_name="step-number"),
                    rx.el.h3("Connect the dots"),
                    rx.el.p("We map your real experience to what the role needs."),
                    class_name="step",
                ),
                rx.el.article(
                    rx.el.div("03 / DELIVERY", class_name="step-number"),
                    rx.el.h3("Make the introduction"),
                    rx.el.p("Review, copy, or download a polished cover letter."),
                    class_name="step",
                ),
                class_name="steps-grid",
            ),
            id="how-it-works",
            class_name="studio-section",
        ),
        rx.el.section(
            rx.el.h2("About"),
            rx.el.p(
                "Cover Letter Studio turns your existing experience into a focused, role-specific introduction. "
                "Every draft is grounded in the details you provide, so the result stays credible and personal.",
                class_name="about-text",
            ),
            id="about",
            class_name="studio-section",
        ),
    )


def index() -> rx.Component:
    return rx.el.main(
        rx.el.div(class_name="page-background"),
        rx.el.div(class_name="page-scrim"),
        nav(),
        rx.el.div(hero(), workspace(), sections(), class_name="page-container"),
        class_name="app",
    )


app = rx.App(
    theme=rx.theme(
        appearance="dark",
        accent_color="orange",
        gray_color="sand",
        radius="large",
        has_background=False,
    ),
    stylesheets=["/styles.css"],
)
app.add_page(index, route="/", title="Cover Letter Studio")
