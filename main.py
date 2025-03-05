import google.generativeai as genai
import time
import os
import logging
from dotenv import load_dotenv
import pypandoc

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load environment variables
load_dotenv()

# --- Configure your API Key ---
GOOGLE_API_KEY = os.environ.get('GOOGLE_API_KEY')
if not GOOGLE_API_KEY:
    raise ValueError("Please set the GOOGLE_API_KEY environment variable in your .env file.")
genai.configure(api_key=GOOGLE_API_KEY)

# --- Book Structure Definition ---
# Now allows chapters without sections
book_structure = {
    "Acknowledgements": [],
    "Foreword: The Veil of Ignorance": [],
    "The Origin of the Saurian Lineage": [
        "The Primordial Genesis: Extraterrestrial Seeding",
        "The Evolutionary Divergence: Terraforming and Adaptation",
        "The First Colonies: Pre-Human Civilizations",
        "The Cataclysmic Shift: Subterranean Migration"
    ],
    "The Xenoarchaeological Evidence: Recovered Artifacts": [
        "The Ancient Texts: Deciphered Glyphs",
        "The Technological Relics: Reverse Engineering",
        "The Biological Specimens: Genetic Analysis",
        "The Architectural Structures: Subterranean Excavations"
    ],
    "The Reptilian Physiology and Psychology": [
        "The Adaptive Morphology: Cold-Blooded Mastery",
        "The Neural Architecture: Cognitive Superiority",
        "The Societal Structure: Hierarchical Imperatives",
        "The Psionic Capabilities: Hidden Influence"
    ],
    "The Genetic Engineering: Enhancing the Lineage": [
        "The Hybridization Programs: Creating Intermediate Forms",
        "The Longevity Protocols: Extending Lifespan",
        "The Cognitive Augmentation: Mental Enhancement",
        "The Physical Adaptation: Environmental Resilience"
    ],
    "The Psionic Network: Collective Consciousness": [
        "The Resonance Frequencies: Interconnected Minds",
        "The Telepathic Communication: Covert Messaging",
        "The Precognitive Abilities: Future Prediction",
        "The Mental Domination: Subjugation Techniques"
    ],
    "The Subterranean Ecosystem: Sustaining the Hidden World": [
        "The Geothermal Energy: Power Sources",
        "The Hydroponic Cultivation: Food Production",
        "The Environmental Control: Atmospheric Regulation",
        "The Advanced Infrastructure: Transportation Networks"
    ],
    "The Historical Interventions: Shaping Human Destiny": [
        "The Ancient Empires: Patronage and Control",
        "The Religious Manipulations: Establishing Dogma",
        "The Technological Advancements: Directed Evolution",
        "The Geopolitical Strategies: Engineered Conflicts"
    ],
    "The Concealment Protocol: Maintaining the Veil": [
        "The Infiltration Strategy: Cultural Mimicry",
        "The Information Control: Media and Propaganda",
        "The Genetic Camouflage: Human Disguise",
        "The Subterranean Networks: Hidden Cities"
    ],
    "The Religious Syncretism: Unified Beliefs": [
        "The Cult of the Serpent: Ancient Worship",
        "The New Age Movements: Controlled Spirituality",
        "The Esoteric Doctrines: Hidden Knowledge",
        "The Messianic Figures: Engineered Saviors"
    ],
    "The Ritualistic Practices: Maintaining Order": [
        "The Symbolic Ceremonies: Reinforcing Allegiance",
        "The Energy Harvesting: Psionic Amplification",
        "The Bloodline Purification: Genetic Integrity",
        "The Prophetic Visions: Interpreting Future Events"
    ],
    "The Elite Symbiosis: The Inner Circle": [
        "The Bloodlines of Power: Inherited Influence",
        "The Secret Societies: Ritual and Allegiance",
        "The Corporate Conglomerates: Resource Acquisition",
        "The Military Industrial Complex: Strategic Supremacy"
    ],
    "The Modern Paradigm: Control in the Digital Age": [
        "The Financial Institutions: Economic Dominance",
        "The Political Puppetry: Global Governance",
        "The Technological Surveillance: Information Gathering",
        "The Pharmaceutical Manipulation: Biological Control"
    ],
    "The Media Manipulation: Shaping Perceptions": [
        "The Mainstream Narratives: Controlling Information",
        "The Entertainment Industry: Subliminal Messaging",
        "The Social Media Platforms: Algorithmic Control",
        "The Educational Institutions: Indoctrination Programs"
    ],
    "The Economic Domination: Global Markets": [
        "The Shadow Banking Systems: Unregulated Finance",
        "The Resource Cartels: Controlling Commodities",
        "The Intellectual Property: Patent Manipulation",
        "The Digital Currency: Financial Control"
    ],
    "The Diplomatic Maneuvers: International Relations": [
        "The Covert Treaties: Secret Agreements",
        "The Geopolitical Strategies: Power Balancing",
        "The Humanitarian Interventions: Engineered Crises",
        "The Global Summits: Orchestrated Meetings"
    ],
    "The Military Applications: Advanced Warfare": [
        "The Directed Energy Weapons: Silent Destruction",
        "The Psionic Warfare: Mental Subjugation",
        "The Biological Weapons: Genetic Targeting",
        "The Space-Based Platforms: Orbital Supremacy"
    ],
    "The Counter-Intelligence: Identifying Resistance": [
        "The Surveillance Networks: Monitoring Activities",
        "The Psychological Profiling: Behavioral Analysis",
        "The Disinformation Campaigns: Spreading Confusion",
        "The Elimination Protocols: Targeted Removal"
    ],
    "The Future Trajectory: Shaping the Next Era": [
        "The Population Control: Resource Management",
        "The Environmental Restructuring: Planetary Modification",
        "The Technological Singularity: Artificial Intelligence",
        "The Transhuman Integration: Merging with Humanity"
    ],
    "The Cosmic Agenda: Interstellar Influence": [
        "The Galactic Alliances: Extraterrestrial Contact",
        "The Resource Acquisition: Planetary Mining",
        "The Strategic Positioning: Interstellar Conflicts",
        "The Temporal Anomalies: Manipulating Time"
    ],
    "The Temporal Control: Altering Timelines": [
        "The Chronal Technologies: Time Manipulation",
        "The Butterfly Effect: Minor Alterations",
        "The Predetermined Outcomes: Engineered Futures",
        "The Multiverse Theory: Parallel Realities"
    ],
    "Appendices": [
        "Chronological Timeline of Reptilian Influence",
        "Glossary of Reptilian Terminology",
        "Notable Reptilian Bloodlines and Affiliates",
        "Recovered Artifact Catalog",
        "Psionic Frequency Charts and Data",
        "Subterranean Network Maps"
    ]
}

# --- Model Configuration ---
model = genai.GenerativeModel('gemini-2.0-flash')
def generate_prompt(chapter_title, section_title=None, previous_content=""):
    """Generates a detailed prompt.  Handles chapters with and without sections."""
    if section_title:
        prompt = f"""
        You are writing a book about reptilian. This is a science fiction project.
        In this fantasy world, the reptilians are real. They live amongst humanity.
        However most human do not know about them. Only certain elites in the society
        know about them, or the people who somehow managed to access those hidden 
        top secrets.  The target audience is for people who are new to the elite 
        and secret world - they need to get up-to-date and comprehensive on how the
        Reptilians interacted with Earth's history and humanity's history. We need
        to discuss how the Reptilians came about, how they are hiding from humanity,
        why they are hiding from humanity, how they have affected humanity, how they
        are affecting humanity, and how they might affect humanity in the future.
        This book is a fictional textbook. This book is written in a way that sounds
        like objective truth. The wordings in the book must not sound like speculation.
        Everything stated in the book has to be written in a way so it sounds like 
        scientific fact. These are done in order to make this book sound like a 
        textbook that you can pick up in a secret restricted library in this 
        fictional universe.
        Chapter: {chapter_title}
        Section: {section_title}
        Previously generated content: {previous_content}
        Write the complete content for the current section ({section_title}).
        Your output must also be extraordinarily long, containing meticulous details.
        Please maintain an academic tone.
        Do not include chapter or section headers.
        """
    else:  # Chapter without sections
        prompt = f"""
        You are writing a book about reptilian. This is a science fiction project.
        In this fantasy world, the reptilians are real. They live amongst humanity.
        However most human do not know about them. Only certain elites in the society
        know about them, or the people who somehow managed to access those hidden 
        top secrets.  The target audience is for people who are new to the elite 
        and secret world - they need to get up-to-date and comprehensive on how the
        Reptilians interacted with Earth's history and humanity's history. We need
        to discuss how the Reptilians came about, how they are hiding from humanity,
        why they are hiding from humanity, how they have affected humanity, how they
        are affecting humanity, and how they might affect humanity in the future.
        This book is a fictional textbook. This book is written in a way that sounds
        like objective truth. The wordings in the book must not sound like speculation.
        Everything stated in the book has to be written in a way so it sounds like 
        scientific fact. These are done in order to make this book sound like a 
        textbook that you can pick up in a secret restricted library in this 
        fictional universe.
        Write the complete content for an entire chapter.
        Chapter: {chapter_title}
        Previously generated content: {previous_content}
        Write the complete content for the chapter ({chapter_title}).
        Your output must also be extraordinarily long, containing meticulous details.
        Please maintain an academic tone.
        Do not include a chapter header.
        """
    return prompt.strip() #Remove leading/trailing whitespaces

def generate_content(chapter_title, section_title=None, previous_content="", max_retries=5, retry_delay=10):
    """Generates content for a chapter or section, handling errors."""
    prompt = generate_prompt(chapter_title, section_title, previous_content)
    for attempt in range(max_retries):
        try:
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            logging.warning(f"Error (attempt {attempt + 1}/{max_retries}): {e}")
            if "429" in str(e):
                logging.warning(f"Rate limit exceeded. Waiting...")
                time.sleep(retry_delay)
            elif attempt < max_retries - 1:
                time.sleep(retry_delay)
            else:
                logging.error(f"Max retries reached.  Could not generate content.")
                return None

def create_markdown(book_content, filename="generative_ai_book.md"):
    """Creates a Markdown file."""
    with open(filename, "w", encoding="utf-8") as f:
        for chapter_title, chapter_content in book_content.items():
            f.write(f"# {chapter_title}\n\n")
            if isinstance(chapter_content, dict):  # Check if it's a dictionary (has sections)
                for section_title, section_content in chapter_content.items():
                    f.write(f"## {section_title}\n\n")
                    f.write(section_content)
                    f.write("\n\n")
            else:  # It's a string (chapter content directly)
                f.write(chapter_content)
                f.write("\n\n")
    logging.info(f"Markdown book saved to {filename}")


def convert_markdown_to_docx_with_pypandoc(markdown_file, docx_file):
    """Converts a Markdown file to DOCX using pypandoc."""
    try:
        extra_args = [
            "--toc",  # Only include the table of contents
        ]
        output = pypandoc.convert_file(markdown_file, 'docx', outputfile=docx_file, extra_args=extra_args, format='markdown')
        if output == "":
             logging.info(f"Successfully converted {markdown_file} to {docx_file}")
        else:
            logging.warning(f"pypandoc conversion may have warnings: {output}")
    except RuntimeError as e:
        logging.error(f"pypandoc conversion failed: {e}")
    except FileNotFoundError:
        logging.error("Pandoc not found.  Make sure it's installed and in your PATH.")
    except OSError as e:
        logging.error(f"An OS error occurred: {e}")
def main():
    """Main function."""
    book_content = {}
    all_previous_content = ""

    for chapter_title, sections in book_structure.items():
        if sections:  # Chapter with sections
            book_content[chapter_title] = {}
            for section_title in sections:
                logging.info(f"Generating: {chapter_title} - {section_title}")
                section_content = generate_content(chapter_title, section_title, all_previous_content)
                if section_content:
                    book_content[chapter_title][section_title] = section_content
                    all_previous_content += f"\n\n{chapter_title} - {section_title}:\n{section_content}"
                else:
                    logging.warning(f"Skipping section: {chapter_title} - {section_title}")
        else:  # Chapter without sections
            logging.info(f"Generating: {chapter_title}")
            chapter_content = generate_content(chapter_title, previous_content=all_previous_content) # No section_title
            if chapter_content:
                book_content[chapter_title] = chapter_content # Store directly, not in a nested dict
                all_previous_content += f"\n\n{chapter_title}:\n{chapter_content}"
            else:
                logging.warning(f"Skipping chapter: {chapter_title}")

    markdown_filename = "generative_ai_book.md"
    docx_filename = "generative_ai_book.docx"
    create_markdown(book_content, markdown_filename)
    convert_markdown_to_docx_with_pypandoc(markdown_filename, docx_filename)

if __name__ == "__main__":
    main()