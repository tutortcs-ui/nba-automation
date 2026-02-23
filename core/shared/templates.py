# =============================================================================
# core/shared/templates.py
# PURPOSE: Generate "Action Taken" sentences for the ATR table.
#          Handles both Course-Exit (High/Mod/Low) and
#          Faculty feedback (Strongly Agree / Agree / Neutral / Disagree).
#
# STRUCTURE per question:
#   Sentence 1 (Data):   "X% of students strongly agreed / highly agreed that..."
#   Sentence 2 (Action): Tiered recommendation based on top positive %
# =============================================================================

import re


# -----------------------------------------------------------------------
# RECOMMENDATION TEMPLATES — three tiers, used by both feedback types
# -----------------------------------------------------------------------

HIGH_TEMPLATES = [
    "Continue encouraging students to explore {topic} through advanced discussions and applied exercises.",
    "Sustain the current teaching approach and challenge students further with real-world applications of {topic}.",
    "Reinforce understanding of {topic} by introducing higher-order thinking tasks and peer-led activities.",
]

MODERATE_TEMPLATES = [
    "Use visual aids, structured examples, and group discussions to help students strengthen their grasp of {topic}.",
    "Incorporate hands-on exercises and case studies to deepen student understanding of {topic}.",
    "Provide supplementary resources and guided practice sessions to consolidate knowledge of {topic}.",
]

LOW_TEMPLATES = [
    "Revisit core concepts of {topic} through remedial sessions and step-by-step guided instruction.",
    "Provide additional support and foundational review of {topic}, with frequent formative assessments to track progress.",
    "Restructure delivery of {topic} using simpler explanations, worked examples, and increased one-on-one support.",
]

# -----------------------------------------------------------------------
# Faculty-specific recommendation templates
# These are tuned to teaching quality statements rather than course content
# -----------------------------------------------------------------------

FACULTY_HIGH_TEMPLATES = [
    "Maintain the current standard of {topic} as students have responded positively.",
    "Continue the effective approach to {topic} and explore further enhancements.",
    "Sustain and document the current best practices related to {topic} for institutional benefit.",
]

FACULTY_MODERATE_TEMPLATES = [
    "Explore additional strategies to enhance {topic} and improve student satisfaction.",
    "Review current methods for {topic} and incorporate student suggestions for improvement.",
    "Consider peer review and collaborative planning to strengthen {topic}.",
]

FACULTY_LOW_TEMPLATES = [
    "Immediate review and corrective action required for {topic}.",
    "Seek mentoring and training opportunities to address gaps in {topic}.",
    "Develop an improvement plan for {topic} with support from the department.",
]


def _pick_template(templates: list, index: int) -> str:
    """Rotate through templates by index so consecutive questions vary."""
    return templates[index % len(templates)]


def _clean_topic(description: str) -> str:
    """
    Trim a question description down to a short topic phrase.
    Strips 'Students will [verb]' prefix if present.
    Used in recommendation sentences as {topic}.
    """
    text = description.strip().rstrip(".")
    text = re.sub(
        r"^students\s+will\s+(understand|learn|grasp|gain\s+insight\s+into|"
        r"explore|develop|apply|analyze|evaluate)\s+",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()
    return text


def _co_verb_phrase(description: str) -> str:
    """
    Convert 'Students will understand X' → 'understood X' for data sentence.
    Falls back to lowercase description if pattern not matched.
    """
    verb_map = {
        "understand": "understood", "learn": "learned", "grasp": "grasped",
        "gain": "gained", "apply": "applied", "analyze": "analyzed",
        "develop": "developed", "explore": "explored", "evaluate": "evaluated",
        "implement": "implemented", "design": "designed", "create": "created",
        "use": "used", "identify": "identified",
    }
    match = re.match(r"^students\s+will\s+(\w+)\s+(.*)", description.strip(), re.IGNORECASE)
    if match:
        verb = match.group(1).lower()
        rest = match.group(2).strip().rstrip(".")
        return f"{verb_map.get(verb, verb + 'ed')} {rest}"
    return description.lower().rstrip(".")


# -----------------------------------------------------------------------
# PUBLIC FUNCTIONS
# -----------------------------------------------------------------------

def generate_course_exit_action_taken(question: dict, index: int = 0) -> str:
    """
    Generate Action Taken text for one Course-Exit CO.

    Data sentence:    "[X]% of students highly agreed that they [verb phrase]."
    Action sentence:  Tiered recommendation based on high_pct.

    Args:
        question: dict with keys: description, high_pct, moderate_pct, low_pct
        index:    position in list (for template variety)
    """
    high_pct    = question["high_pct"]
    description = question["description"]

    data_sentence = (
        f"{high_pct}% of students highly agreed that they "
        f"{_co_verb_phrase(description)}."
    )

    topic = _clean_topic(description)
    if high_pct >= 75:
        template = _pick_template(HIGH_TEMPLATES, index)
    elif high_pct >= 50:
        template = _pick_template(MODERATE_TEMPLATES, index)
    else:
        template = _pick_template(LOW_TEMPLATES, index)

    return f"{data_sentence} {template.format(topic=topic)}"


def generate_faculty_action_taken(question: dict, index: int = 0) -> str:
    """
    Generate Action Taken text for one Faculty feedback question.

    Data sentence:    "[X]% of students strongly agreed that [question text]."
    Action sentence:  Tiered recommendation based on strongly_agree_pct.

    Args:
        question: dict with keys: description, strongly_agree_pct,
                  agree_pct, neutral_pct, disagree_pct
        index:    position in list (for template variety)
    """
    sa_pct      = question.get("strongly_agree_pct", 0)
    description = question["description"].rstrip(".")

    data_sentence = (
        f"{sa_pct}% of students strongly agreed that {description.lower()}."
    )

    topic = description.lower()
    if sa_pct >= 75:
        template = _pick_template(FACULTY_HIGH_TEMPLATES, index)
    elif sa_pct >= 50:
        template = _pick_template(FACULTY_MODERATE_TEMPLATES, index)
    else:
        template = _pick_template(FACULTY_LOW_TEMPLATES, index)

    return f"{data_sentence} {template.format(topic=topic)}"


def generate_all_course_exit_action_taken(question_data: list) -> list:
    """Generate Action Taken for all Course-Exit COs. Returns list of strings."""
    return [generate_course_exit_action_taken(q, i) for i, q in enumerate(question_data)]


def generate_all_faculty_action_taken(question_data: list) -> list:
    """Generate Action Taken for all Faculty questions. Returns list of strings."""
    return [generate_faculty_action_taken(q, i) for i, q in enumerate(question_data)]
