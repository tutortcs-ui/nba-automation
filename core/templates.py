# =============================================================================
# templates.py
# PURPOSE: Generate the two sentences in the ATR "Action Taken" column
#          using rule-based templates (no AI API needed).
#
# SENTENCE 1 (Data sentence):
#   Standardized format: "[X]% of students highly agreed that they [CO verb phrase]."
#
# SENTENCE 2 (Recommendation sentence):
#   Tone changes based on High%:
#     >= 75%  → Positive reinforcement (continue current approach)
#     50-74%  → Strengthen with methods (visual aids, discussions, case studies)
#     < 50%   → Remedial action (revisit core concepts, additional support)
# =============================================================================


# -----------------------------------------------------------------------
# RECOMMENDATION TEMPLATES
# Each template uses {topic} as a placeholder for the CO description text.
# Three tiers based on High% score.
# -----------------------------------------------------------------------

HIGH_TEMPLATES = [
    # High% >= 75 — students are doing well, reinforce and deepen
    "Continue encouraging students to explore {topic} through advanced discussions and applied exercises.",
    "Sustain the current teaching approach and challenge students further with real-world applications of {topic}.",
    "Reinforce understanding of {topic} by introducing higher-order thinking tasks and peer-led activities.",
]

MODERATE_TEMPLATES = [
    # Moderate% dominant or High% 50-74 — good but needs strengthening
    "Use visual aids, structured examples, and group discussions to help students strengthen their grasp of {topic}.",
    "Incorporate hands-on exercises and case studies to deepen student understanding of {topic}.",
    "Provide supplementary resources and guided practice sessions to consolidate knowledge of {topic}.",
]

LOW_TEMPLATES = [
    # High% < 50 — significant gap, remedial action needed
    "Revisit core concepts of {topic} through remedial sessions and step-by-step guided instruction.",
    "Provide additional support and foundational review of {topic}, with frequent formative assessments to track progress.",
    "Restructure delivery of {topic} using simpler explanations, worked examples, and increased one-on-one support.",
]


def _pick_template(templates: list[str], co_index: int) -> str:
    """
    Pick a template from a list using the CO index.
    This ensures different COs get different recommendation sentences
    even when they fall in the same tier — adds variety to the document.
    """
    return templates[co_index % len(templates)]


def _extract_topic_phrase(description: str) -> str:
    """
    Clean up the CO description to use as the {topic} placeholder.
    Strips leading/trailing whitespace and newlines.
    Converts to lowercase for mid-sentence use.
    
    E.g.: "Students will understand various learning paradigms..."
       -> "various learning paradigms and foundational techniques in deep learning frameworks"
    
    We strip the "Students will [verb]" prefix since the data sentence
    already uses that, and the recommendation sentence needs just the topic.
    """
    text = description.strip()

    # Remove common prefixes like "Students will understand", "Students will learn", etc.
    import re
    text = re.sub(
        r"^students\s+will\s+(understand|learn|grasp|gain insight into|explore|develop|apply|analyze|evaluate)\s+",
        "",
        text,
        flags=re.IGNORECASE
    ).strip()

    # Remove trailing period if present
    text = text.rstrip(".")

    return text


def _extract_verb_phrase(description: str) -> str:
    """
    Extract the verb phrase from CO description for the data sentence.
    
    Data sentence format:
    "[X]% of students highly agreed that they [verb phrase]."
    
    E.g.: "Students will understand various learning paradigms..."
       -> "understood various learning paradigms..."
    
    We try to convert present tense to past tense for natural reading.
    """
    text = description.strip()

    # Map common CO verbs to their past-tense equivalents
    verb_map = {
        "understand":  "understood",
        "learn":       "learned",
        "grasp":       "grasped",
        "gain":        "gained",
        "apply":       "applied",
        "analyze":     "analyzed",
        "develop":     "developed",
        "explore":     "explored",
        "evaluate":    "evaluated",
        "implement":   "implemented",
        "design":      "designed",
        "create":      "created",
        "use":         "used",
        "identify":    "identified",
    }

    import re

    # Replace "Students will [verb]" with just the past-tense verb
    match = re.match(
        r"^students\s+will\s+(\w+)\s+(.*)",
        text,
        flags=re.IGNORECASE
    )

    if match:
        verb = match.group(1).lower()
        rest = match.group(2).strip().rstrip(".")
        past_verb = verb_map.get(verb, verb + "ed")  # fallback: add 'ed'
        return f"{past_verb} {rest}"

    # Fallback: return as-is (lowercase)
    return text.lower().rstrip(".")


def generate_action_taken(co: dict, co_index: int = 0) -> str:
    """
    MAIN ENTRY POINT for this module.
    
    Generate the full "Action Taken" text for one CO.
    Returns two sentences joined together.
    
    Args:
        co:       dict with keys: code, description, high_pct, moderate_pct, low_pct
        co_index: position of this CO in the list (0-based) — used for template variety
    
    Returns:
        String with two sentences:
        "89% of students highly agreed that they understood... 
         Continue encouraging students to explore..."
    """
    high_pct = co["high_pct"]
    description = co["description"]

    # --- SENTENCE 1: Data sentence ---
    verb_phrase = _extract_verb_phrase(description)
    data_sentence = f"{high_pct}% of students highly agreed that they {verb_phrase}."

    # --- SENTENCE 2: Recommendation sentence ---
    topic = _extract_topic_phrase(description)

    if high_pct >= 75:
        template = _pick_template(HIGH_TEMPLATES, co_index)
    elif high_pct >= 50:
        template = _pick_template(MODERATE_TEMPLATES, co_index)
    else:
        template = _pick_template(LOW_TEMPLATES, co_index)

    recommendation_sentence = template.format(topic=topic)

    return f"{data_sentence} {recommendation_sentence}"


def generate_all_action_taken(co_data: list[dict]) -> list[str]:
    """
    Generate Action Taken text for all COs.
    
    Returns list of strings in same order as co_data.
    """
    return [
        generate_action_taken(co, i)
        for i, co in enumerate(co_data)
    ]
