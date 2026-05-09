---
name: planificacion-profesor
description: Genera planificaciones semanales completas para profesores de cualquier asignatura y nivel educativo. Exporta a Word (.docx). Usa /planificacion para iniciar.
---

# Skill: Generador de Planificaciones Semanales

Cuando el usuario ejecuta `/planificacion`, sigue este flujo exacto.

---

## PASO 1 — Verificar contexto de la profesora

Busca el archivo `mi-contexto.md` en el directorio de trabajo actual.

**Si NO existe:** dile al usuario:

> No encontré tu archivo de contexto `mi-contexto.md`. Este archivo guarda tu información permanente para que no tengas que repetirla cada semana.
>
> Voy a hacerte algunas preguntas para crearlo. Solo las respondes una vez.

Luego hazle estas preguntas (todas juntas en un solo mensaje):

```
Para crear tu perfil de profesora necesito saber:

1. ¿Cuál es tu nombre completo?
2. ¿En qué institución / escuela trabajas?
3. ¿Qué asignatura o área enseñas? (Ej: Lenguaje, Matemáticas, Ciencias, Ed. Parvularia, etc.)
4. ¿Qué nivel o curso tienes a cargo? (Ej: NT2, 3° básico, 1° medio, etc.)
5. ¿Cuántos estudiantes tiene tu curso?
6. ¿Hay perfiles o grupos diferenciados en tu curso? (Ej: niños con NEE, niveles A/B/C, etc.)
7. ¿Qué marco curricular usas? (Ej: BCEP, Bases Curriculares MINEDUC, programa propio, etc.)
8. ¿Cómo debe estructurarse tu planificación? (Ej: por bloques horarios, por objetivos, con tabla de OA, libre, etc.)
9. ¿Hay algo más sobre tu escuela o grupo que deba saber para generar buenas planificaciones?
```

Cuando el usuario responda, crea el archivo `mi-contexto.md` en el directorio actual con toda esa información organizada, y confirma que quedó guardado.

**Si SÍ existe:** léelo silenciosamente y continúa al Paso 2.

---

## PASO 2 — Preguntas de la semana

Hazle estas preguntas en un solo mensaje:

```
Perfecto. Ahora cuéntame sobre esta semana:

1. ¿Cuáles son las fechas de la semana? (ej: 12 al 16 de mayo)
2. ¿Cuál es la unidad o temática que estás trabajando este mes?
3. ¿Cuál es el foco específico de esta semana? (contenido, habilidad u objetivo que quieras trabajar)
4. ¿Hay efemérides, eventos escolares o días sin clases esta semana?
5. ¿Tienes ideas propias para alguna actividad o día en particular?
6. ¿Hay algo más que quieras agregar o que deba considerar para esta semana?
```

Espera la respuesta antes de continuar.

---

## PASO 3 — Generar la planificación

Con la información del contexto + las respuestas de la semana, genera la planificación completa.

### Reglas de generación (obligatorias)

- **Estructura:** usa la estructura que la profesora indicó en su contexto. Si no especificó una, usa: Cabecera → Resumen semanal (tabla) → Fichas detalladas por bloque/sesión.
- **OA / objetivos:** cópialos LITERAL del marco curricular indicado. No reformules ni parafrasees.
- **Diferenciación:** si hay perfiles o grupos, diferencia las actividades por perfil en cada sesión.
- **Actividades:** estructura cada sesión con Inicio / Desarrollo / Cierre explícitos.
- **Recursos:** concretos, de bajo costo, disponibles en un aula típica.
- **Tono:** mezclado (técnico pero cercano). Sin "el/la docente" — usar primera persona o nombre de los estudiantes.
- **Longitud:** completa pero concisa. Frases directas, sin relleno.
- **Adecuaciones:** incorpora las necesidades especiales del grupo si las hay.

### Formato de salida

Genera la planificación completa en Markdown con:

1. **Tabla de cabecera** (Mes · Nivel · Educadora · Temática · Semana)
2. **Tabla de resumen semanal** (días × bloques/sesiones con horario y tipo)
3. **Ficha detallada** de cada sesión/bloque con todos los campos

Muestra la planificación completa en pantalla para que el usuario la revise.

---

## PASO 4 — Refinamiento

Después de mostrar la planificación, pregunta:

> ¿Quieres cambiar algo? Puedes pedirme ajustes específicos (ej: "cambia la actividad del miércoles", "el cierre del martes es muy corto", "agrega un indicador más al bloque 2 del lunes") o decirme que está lista para exportar.

Aplica los cambios que pida y muestra la versión actualizada. Repite hasta que el usuario confirme que está conforme.

---

## PASO 5 — Exportar a Word (.docx)

Cuando el usuario confirme que la planificación está lista, genera el archivo Word.

### Preparar el JSON

Convierte la planificación a JSON con esta estructura:

```json
{
  "cabecera": {
    "mes": "",
    "nivel": "",
    "educadora": "",
    "tematica": "",
    "semana": ""
  },
  "resumen": {
    "lunes": {
      "emoji": "📅",
      "fecha": "Lunes DD",
      "bloques": [
        {"numero": 1, "horario": "HH:MM–HH:MM", "tipo": "Nombre del bloque", "tema": "Tema"}
      ]
    }
  },
  "dias": [
    {
      "dia": "LUNES DD DE MES",
      "bloques": [
        {
          "numero": 1,
          "horario": "HH:MM a HH:MM",
          "tipo": "Nombre del bloque",
          "meta": "Objetivo específico del bloque.",
          "ambito_nucleo": "Área / Núcleo o Eje",
          "oa": "OA X: texto literal del objetivo",
          "oat": "Objetivo transversal si aplica",
          "habilidad": "Verbo1 · Verbo2",
          "inicio": "Descripción del inicio.",
          "desarrollo": "Descripción del desarrollo.",
          "cierre": "Descripción del cierre.",
          "indicadores": ["Indicador 1.", "Indicador 2."],
          "evaluacion_procedimiento": "Observación",
          "evaluacion_instrumento": "Lista de apreciación",
          "recursos": ["Recurso 1", "Recurso 2"],
          "nivel_fonologico": "Si aplica, trabajo fonológico. Si no aplica: —",
          "nivel_semantico": "Vocabulario y trabajo semántico.",
          "nivel_pragmatico": "Uso comunicativo en contexto."
        }
      ]
    }
  ]
}
```

### Generar el .docx

1. Guarda el JSON en `planificacion_SEMANA.json` en el directorio actual.

2. Verifica si `generate_docx.py` existe en el directorio actual. Si no, cópialo desde `~/.claude/skills/planificacion-profesor/generate_docx.py`.

3. Verifica si `python-docx` está instalado:
   ```bash
   python3 -c "import docx" 2>/dev/null || pip3 install python-docx --user -q || pip3 install python-docx --break-system-packages -q
   ```

4. Ejecuta:
   ```bash
   python3 generate_docx.py --input planificacion_SEMANA.json --output planificacion_SEMANA.docx
   ```

5. Confirma al usuario que el archivo quedó en el directorio actual y listo para abrir en Word.

---

## Notas importantes

- Si el usuario dice `/planificacion nueva` o `/planificacion reset`, ignora el `mi-contexto.md` existente y empieza desde el Paso 1 para actualizar el perfil.
- Si el usuario dice `/planificacion revisar`, lee la planificación más reciente en el directorio y aplica el checklist del contexto para sugerir mejoras.
- Siempre responde en español.
- Nunca inventes OA o referencias curriculares — si no sabes el número exacto, indica claramente que el usuario debe verificarlo.
