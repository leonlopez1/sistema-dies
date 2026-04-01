#!/usr/bin/env python3
"""
Generador de Contenido — Consultoría Estratégica para PYMEs Colombianas
Uso: python content_generator.py [tarea] [opciones]
"""

import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

import anthropic

# ── Configuración del negocio ────────────────────────────────────────────────

BUSINESS_CONTEXT = """
## CONTEXTO DEL NEGOCIO

**Servicio:** Consultoría Estratégica para PYMEs colombianas (10-30 empleados)
**Sectores target:** Centros estéticos, clínicas odontológicas, empresas de manufactura/maquila,
  comercio/venta de productos físicos, producción de alimentos, metalmecánica
**Ciudades:** Bogotá, Medellín, Cali, Barranquilla, Bucaramanga
**Propuesta de valor:** Ayudo a dueños de PYMEs a estructurar su operación para que crezcan
  sin depender 100% del fundador — con sistemas, KPIs y procesos que funcionan
**Diferenciador:** No es consultoría genérica. Trabajo con industrias tradicionales que no
  tienen tiempo para teoría — solo resultados medibles en 90 días
**Paquetes:**
  - Diagnóstico: 1 mes, $3-5M COP
  - Estratégico: 3 meses, $8-12M COP
  - Integral: 6 meses, $15-22M COP
**Tono:** Directo, sin buzzwords, con datos y ejemplos reales. Como un amigo empresario
  que sabe de números, no como un consultor de corbata.

## REGLAS DE ESCRITURA
1. Español colombiano natural (tuteo o ustedeo según contexto)
2. NUNCA usar: sinergia, disruptivo, potenciar, apalancar, robusto, ecosistema, holístico
3. SIEMPRE incluir números y ejemplos concretos
4. LinkedIn hooks deben provocar scroll-stop — primera línea es todo
5. Aplicar psicología (Loss Aversion, Social Proof, Anchoring, Reciprocity) de forma ética
"""

SKILLS_DIR = Path(__file__).parent / "skills"
OUTPUTS_DIR = Path(__file__).parent / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

# ── Carga de skills ──────────────────────────────────────────────────────────

def load_skill(skill_name: str) -> str:
    """Carga el contenido de un archivo de skill."""
    skill_file = SKILLS_DIR / f"{skill_name}.md"
    if not skill_file.exists():
        return f"[Skill '{skill_name}' no encontrada]"
    return skill_file.read_text(encoding="utf-8")


def load_skills(*skill_names: str) -> str:
    """Carga y concatena múltiples skills."""
    parts = []
    for name in skill_names:
        content = load_skill(name)
        parts.append(f"## SKILL CARGADA: {name.upper()}\n\n{content}")
    return "\n\n---\n\n".join(parts)


# ── Cliente Anthropic ────────────────────────────────────────────────────────

def get_client() -> anthropic.Anthropic:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: La variable ANTHROPIC_API_KEY no está configurada.")
        print("Ejecuta: export ANTHROPIC_API_KEY='tu-api-key'")
        sys.exit(1)
    return anthropic.Anthropic(api_key=api_key)


def generate(system_prompt: str, user_prompt: str, task_label: str) -> str:
    """Llama a Claude y retorna el texto generado (con streaming)."""
    client = get_client()

    print(f"\n⚙  Generando: {task_label}...\n{'─' * 60}")

    full_text = ""
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            full_text += text

    print(f"\n{'─' * 60}\n")
    return full_text


def save_output(content: str, task_type: str, sector: str = "") -> Path:
    """Guarda el resultado en el directorio de outputs."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = f"_{sector}" if sector else ""
    filename = f"{task_type}{suffix}_{timestamp}.md"
    output_path = OUTPUTS_DIR / filename
    output_path.write_text(content, encoding="utf-8")
    print(f"✅  Guardado en: {output_path}")
    return output_path


# ── Tareas de generación ─────────────────────────────────────────────────────

def tarea_posts_linkedin(sector: str = "", tema: str = "") -> str:
    """Genera 5 posts semanales para LinkedIn."""
    skills = load_skills("copywriting", "marketing-psychology", "social-content")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    sector_text = f"enfocados en el sector de **{sector}**" if sector else "variando entre los sectores target"
    tema_text = f"con énfasis en el tema: **{tema}**" if tema else ""

    user = f"""Genera 5 posts completos para LinkedIn para esta semana, {sector_text}. {tema_text}

Estructura exacta:
1. **Caso Real** — Hook de curiosidad + situación real anonimizada + resultado medible
2. **Insight de Industria** — Dato del sector colombiano + contexto + implicación práctica
3. **Educativo** — Framework o tip accionable que el lector puede usar hoy
4. **Behind the Scenes** — Historia personal del trabajo de consultoría o aprendizaje
5. **Hot Take** — Opinión contrarian sobre gestión de PYMEs o el sector

Para cada post incluye:
- El texto completo listo para publicar
- Nota breve sobre la psicología aplicada
- 3-5 hashtags relevantes en español

Formato de salida para cada post:
---
## Post [N]: [Tipo]
**Psicología aplicada:** [qué principio y por qué]

[texto completo del post]

[hashtags]
---
"""

    result = generate(system, user, f"5 Posts LinkedIn - {sector or 'multisector'}")
    save_output(result, "linkedin_posts", sector.replace(" ", "_").lower() if sector else "multisector")
    return result


def tarea_secuencia_mensajes(sector: str) -> str:
    """Genera secuencia de 5 toques de LinkedIn para un sector."""
    skills = load_skills("copywriting", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    user = f"""Genera una secuencia completa de 5 mensajes directos de LinkedIn para prospectar dueños de **{sector}** colombianos.

Cada mensaje debe:
- Sonar completamente humano, NO automatizado
- Tener propósito distinto (curiosidad → valor → caso → diagnóstico → seguimiento)
- Ser corto (máximo 100 palabras por mensaje)
- No copiar el tono corporativo

Estructura:
- **Mensaje 1 (Día 1)**: Conexión + curiosidad — referencia algo específico del sector
- **Mensaje 2 (Día 4)**: Valor gratuito — insight o dato relevante para su industria
- **Mensaje 3 (Día 8)**: Caso social proof — resultado anónimo del mismo sector
- **Mensaje 4 (Día 12)**: Diagnóstico gratuito — propuesta de 20 min sin compromiso
- **Mensaje 5 (Día 17)**: Cierre suave — último intento con puerta abierta

Para cada mensaje:
- Texto listo para copiar y pegar
- Nota de personalización (qué adaptar al prospecto específico)
- Objetivo psicológico del mensaje
"""

    result = generate(system, user, f"Secuencia Mensajes - {sector}")
    save_output(result, "secuencia_mensajes", sector.replace(" ", "_").lower())
    return result


def tarea_emails_frios(sector: str) -> str:
    """Genera secuencia de 3 emails fríos para un sector."""
    skills = load_skills("copywriting", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    user = f"""Genera una secuencia de 3 emails fríos para dueños de **{sector}** colombianos.

**Email 1 — Curiosidad (Día 1)**
- Subject line que provoca apertura (máximo 50 caracteres, sin spam words)
- Abre con hook de problema específico del sector
- No vende nada, genera curiosidad

**Email 2 — Valor (Día 4)**
- Subject line de seguimiento
- Entrega insight o mini-framework del sector
- Un solo dato concreto y útil
- CTA suave: responder con una pregunta

**Email 3 — CTA directo (Día 9)**
- Subject line directo
- Referencia los emails anteriores brevemente
- Propone llamada/reunión de 20 min
- Hace fácil el "sí"

Para cada email:
- Subject line + preview text (primeras 50 caracteres del cuerpo)
- Cuerpo completo
- Nota sobre el hook o técnica usada
"""

    result = generate(system, user, f"Emails Fríos - {sector}")
    save_output(result, "emails_frios", sector.replace(" ", "_").lower())
    return result


def tarea_carrusel_instagram(tema: str, sector: str = "") -> str:
    """Genera guión de carrusel de Instagram."""
    skills = load_skills("copywriting", "social-content", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    sector_text = f" para dueños de **{sector}**" if sector else " para dueños de PYMEs colombianas"

    user = f"""Genera el guión completo de un carrusel de Instagram educativo{sector_text} sobre el tema: **{tema}**

El carrusel debe tener 9 slides:
- **Slide 1 (Portada)**: Título provocador (máximo 8 palabras) + subtítulo breve
- **Slide 2**: El problema o situación que reconocen
- **Slides 3-7**: Contenido principal (1 idea concreta por slide)
- **Slide 8**: Resumen de los 3-5 puntos clave
- **Slide 9**: CTA claro + llamado a la acción

Para cada slide incluye:
- **Texto principal**: Lo que va en el slide (máximo 30 palabras, para que sea legible)
- **Texto de apoyo**: Detalle adicional si aplica (máximo 20 palabras)
- **Nota de diseño**: Sugerencia visual o de color

También incluye:
- Caption completo para la publicación (80-150 palabras)
- 10 hashtags relevantes
"""

    result = generate(system, user, f"Carrusel Instagram - {tema}")
    label = tema.replace(" ", "_").lower()[:30]
    save_output(result, "carrusel_instagram", label)
    return result


def tarea_propuesta(empresa: str, sector: str, empleados: int = 15, ciudad: str = "Colombia") -> str:
    """Genera propuesta comercial personalizada."""
    skills = load_skills("copywriting", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    user = f"""Genera una propuesta comercial completa para:
- **Empresa/Sector**: {empresa} — sector: {sector}
- **Tamaño**: ~{empleados} empleados
- **Ciudad**: {ciudad}

La propuesta debe usar estructura **Good / Better / Best**:
- **Good** → Paquete Diagnóstico ($3-5M COP, 1 mes)
- **Better** → Paquete Estratégico ($8-12M COP, 3 meses) ← Recomendado
- **Best** → Paquete Integral ($15-22M COP, 6 meses)

Incluir en la propuesta:
1. **Resumen ejecutivo** (el problema que tienen, en sus palabras)
2. **Por qué esto importa ahora** (loss aversion + costo de la inacción)
3. **Cómo trabajamos** (proceso en 3-4 pasos simples)
4. **Los 3 paquetes** con precios, entregables específicos y resultados esperados
5. **Caso de éxito** (ejemplo anonimizado del mismo sector)
6. **Próximos pasos** (qué pasa después de decir sí)
7. **Garantía o promesa** de resultados medibles

Aplicar: Anchoring (mostrar el mayor primero), Social Proof (caso del sector), Loss Aversion (costo de no hacer nada)
"""

    result = generate(system, user, f"Propuesta Comercial - {sector}")
    label = sector.replace(" ", "_").lower()
    save_output(result, "propuesta_comercial", label)
    return result


def tarea_mejorar_copy(texto: str) -> str:
    """Analiza y mejora un texto existente."""
    skills = load_skills("copy-editing", "copywriting")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    user = f"""Analiza y mejora el siguiente texto de marketing/ventas:

---
{texto}
---

Entrega:
1. **Diagnóstico**: Los 3 problemas principales del texto original
2. **Versión mejorada**: El texto completo reescrito
3. **Lo que cambié y por qué**: Explicación de los cambios clave (máximo 5 puntos)
"""

    result = generate(system, user, "Edición de Copy")
    save_output(result, "copy_mejorado", "")
    return result


def tarea_ideas_contenido(sector: str = "", cantidad: int = 5) -> str:
    """Genera ideas de contenido con hooks listos."""
    skills = load_skills("marketing-ideas", "social-content", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    sector_text = f"del sector **{sector}**" if sector else "variando entre los sectores target"

    user = f"""Genera {cantidad} ideas de contenido {sector_text} para esta semana.

Para cada idea:
- **Título/Hook listo**: La primera línea que usarías (lista para publicar)
- **Plataforma**: LinkedIn o Instagram
- **Tipo**: Caso real / Insight / Educativo / Hot take / Carrusel
- **Por qué funciona**: Qué principio psicológico activa y por qué resonará con el dueño de PYME
- **Ángulo único**: Qué lo hace diferente a lo genérico

Las ideas deben cubrir al menos 3 tipos diferentes de contenido y al menos 2 plataformas.
"""

    result = generate(system, user, f"Ideas de Contenido - {sector or 'multisector'}")
    save_output(result, "ideas_contenido", sector.replace(" ", "_").lower() if sector else "multisector")
    return result


def tarea_script_diagnostico(sector: str) -> str:
    """Genera script de sesión diagnóstico de 20 minutos."""
    skills = load_skills("copywriting", "marketing-psychology")
    system = f"{BUSINESS_CONTEXT}\n\n{skills}"

    user = f"""Genera el script completo de una sesión de diagnóstico de 20 minutos para un dueño de **{sector}**.

El script debe incluir:

**Apertura (2 min)**
- Cómo empezar para generar confianza inmediata
- Qué decir para bajar defensas

**Preguntas diagnósticas por fase (12 min)**
- 3-4 preguntas que descubran el dolor real (no el síntoma)
- Preguntas específicas del sector {sector}
- Cómo escuchar y qué buscar en las respuestas
- Señales de que hay oportunidad real de trabajar juntos

**Espejo del problema (3 min)**
- Cómo resumir lo que escuchaste en sus propias palabras
- Cómo hacer que el prospecto sienta que lo entiendes

**Transición a propuesta (3 min)**
- Cómo pasar naturalmente del diagnóstico a la propuesta
- Qué decir para crear urgencia real (no artificial)
- Cómo cerrar el siguiente paso sin presión

Para cada sección incluye:
- Guión sugerido (entre comillas, listo para practicar)
- Notas del consultor (qué observar, qué evitar)
"""

    result = generate(system, user, f"Script Diagnóstico - {sector}")
    save_output(result, "script_diagnostico", sector.replace(" ", "_").lower())
    return result


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generador de contenido para Consultoría Estratégica — PYMEs Colombia",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
TAREAS DISPONIBLES:
  linkedin    Genera 5 posts semanales para LinkedIn
  mensajes    Genera secuencia de 5 mensajes directos de LinkedIn
  emails      Genera secuencia de 3 emails fríos
  carrusel    Genera guión de carrusel para Instagram
  propuesta   Genera propuesta comercial Good/Better/Best
  editar      Mejora un copy existente
  ideas       Genera ideas de contenido con hooks
  diagnostico Genera script de sesión diagnóstico

EJEMPLOS:
  python content_generator.py linkedin --sector "clínicas estéticas"
  python content_generator.py mensajes --sector "odontología"
  python content_generator.py emails --sector "manufactura"
  python content_generator.py carrusel --tema "5 señales de que tu negocio te controla"
  python content_generator.py propuesta --empresa "Ferretería Construmax" --sector "comercio" --ciudad "Bogotá"
  python content_generator.py ideas --sector "producción de alimentos" --cantidad 7
  python content_generator.py diagnostico --sector "metalmecánica"
  python content_generator.py editar --archivo mi_texto.txt
        """,
    )
    parser.add_argument("tarea", help="Tipo de contenido a generar")
    parser.add_argument("--sector", default="", help="Sector objetivo")
    parser.add_argument("--tema", default="", help="Tema del contenido (para carruseles e ideas)")
    parser.add_argument("--empresa", default="empresa del sector", help="Nombre o descripción de la empresa (para propuestas)")
    parser.add_argument("--ciudad", default="Colombia", help="Ciudad del cliente (para propuestas)")
    parser.add_argument("--empleados", type=int, default=15, help="Número de empleados (para propuestas)")
    parser.add_argument("--cantidad", type=int, default=5, help="Cantidad de ideas a generar")
    parser.add_argument("--archivo", default="", help="Archivo de texto a editar (para tarea 'editar')")
    parser.add_argument("--texto", default="", help="Texto directo a editar (para tarea 'editar')")

    args = parser.parse_args()
    tarea = args.tarea.lower().strip()

    # Mapa de tareas
    if tarea in ("linkedin", "posts", "posts-linkedin"):
        tarea_posts_linkedin(sector=args.sector, tema=args.tema)

    elif tarea in ("mensajes", "secuencia", "secuencia-mensajes"):
        if not args.sector:
            print("ERROR: Especifica el sector con --sector")
            sys.exit(1)
        tarea_secuencia_mensajes(sector=args.sector)

    elif tarea in ("emails", "email", "emails-frios"):
        if not args.sector:
            print("ERROR: Especifica el sector con --sector")
            sys.exit(1)
        tarea_emails_frios(sector=args.sector)

    elif tarea in ("carrusel", "instagram", "carousel"):
        tema = args.tema or "Cómo estructurar una PYME para crecer sin depender del fundador"
        tarea_carrusel_instagram(tema=tema, sector=args.sector)

    elif tarea in ("propuesta", "propuesta-comercial"):
        if not args.sector:
            print("ERROR: Especifica el sector con --sector")
            sys.exit(1)
        tarea_propuesta(
            empresa=args.empresa,
            sector=args.sector,
            empleados=args.empleados,
            ciudad=args.ciudad,
        )

    elif tarea in ("editar", "edit", "mejorar", "copy-editing"):
        texto = args.texto
        if not texto and args.archivo:
            archivo_path = Path(args.archivo)
            if not archivo_path.exists():
                print(f"ERROR: El archivo '{args.archivo}' no existe.")
                sys.exit(1)
            texto = archivo_path.read_text(encoding="utf-8")
        if not texto:
            print("Pega el texto a editar (termina con una línea que solo diga 'EOF'):")
            lineas = []
            while True:
                linea = input()
                if linea.strip() == "EOF":
                    break
                lineas.append(linea)
            texto = "\n".join(lineas)
        tarea_mejorar_copy(texto=texto)

    elif tarea in ("ideas", "ideas-contenido", "content-ideas"):
        tarea_ideas_contenido(sector=args.sector, cantidad=args.cantidad)

    elif tarea in ("diagnostico", "script", "script-diagnostico"):
        if not args.sector:
            print("ERROR: Especifica el sector con --sector")
            sys.exit(1)
        tarea_script_diagnostico(sector=args.sector)

    else:
        print(f"ERROR: Tarea '{tarea}' no reconocida.")
        print("Tareas válidas: linkedin, mensajes, emails, carrusel, propuesta, editar, ideas, diagnostico")
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
