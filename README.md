# Skill: Generador de Planificaciones para Profesores

Skill para Claude Code que genera planificaciones semanales completas en Word (.docx) para profesores de cualquier asignatura y nivel educativo.

Creado por [@sebarobles](https://github.com/robles311)

---

## ¿Qué hace?

- Guarda tu perfil de profesora una sola vez (`mi-contexto.md`)
- Cada semana solo le dices la temática y el foco — el resto lo recuerda
- Genera la planificación completa con estructura pedagógica
- Exporta directamente a Word (.docx) listo para entregar

## Requisitos

- [Claude Code](https://claude.ai/code) instalado (Mac y Windows)
- Suscripción **Claude Pro** (no necesita API key separada)
- Python 3 instalado ([Mac](https://www.python.org/downloads/) · [Windows](https://www.python.org/downloads/windows/))

## Instalación

```bash
npx skills add robles311/planificacion-profesor -g -y
```

O manualmente: clona este repo dentro de `~/.claude/skills/planificacion-profesor/`

## Uso

1. Crea una carpeta para tus planificaciones (solo la primera vez):
   ```bash
   mkdir ~/Documents/mis-planificaciones
   ```

2. Abre Claude Code en esa carpeta y escribe:
   ```
   /planificacion
   ```

3. La primera vez te pide tu perfil (nombre, escuela, asignatura, nivel, etc.) y lo guarda para siempre.

4. Cada semana solo respondes 6 preguntas rápidas sobre la semana.

5. Cuando está lista, se exporta automáticamente a `.docx`.

## Comandos disponibles

| Comando | Qué hace |
|---|---|
| `/planificacion` | Genera la planificación de la semana |
| `/planificacion nueva` | Actualiza tu perfil de profesora |
| `/planificacion revisar` | Revisa y mejora la planificación más reciente |

## Funciona para cualquier asignatura

Lenguaje · Matemáticas · Ciencias · Historia · Ed. Física · Ed. Parvularia · y cualquier otra.

---

Hecho con ❤️ para hacer más fácil la vida de los profesores.
