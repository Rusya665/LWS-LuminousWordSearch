"""
lws_engine.thesaurus
====================
Provides scientific synonym expansion and general English WordNet thesaurus lookups
for LWS (Luminous Word Search).
"""

import logging
import re
from typing import Dict, List, Optional, Set

logger = logging.getLogger(__name__)

SCIENTIFIC_SYNONYMS: Dict[str, List[str]] = {
    # Materials & Polymers
    "pdms": ["polydimethylsiloxane", "silicone", "dimethicone", "sylgard", "siloxane"],
    "silicone": ["pdms", "polydimethylsiloxane", "siloxane", "organosilicon"],
    "gelma": ["gelatin methacryloyl", "gelatin methacrylate", "gelatin-methacrylate"],
    "udma": ["urethane dimethacrylate", "diurethane dimethacrylate"],
    "pegda": ["poly(ethylene glycol) diacrylate", "polyethylene glycol diacrylate"],
    "tegda": ["triethylene glycol diacrylate"],
    "cellulose": ["cncs", "cellulose nanocrystals", "cnfs", "nanocellulose"],
    "hydrogel": ["microgel", "polymeric network", "crosslinked network"],

    # Photoinitiators & Inhibitors
    "tpo": ["diphenyl(2,4,6-trimethylbenzoyl)phosphine oxide", "darocur tpo", "lucirin tpo"],
    "tpo-l": ["ethyl (2,4,6-trimethylbenzoyl) phenylphosphinate", "ethyl phenyl(2,4,6-trimethylbenzoyl)phosphinate"],
    "bapo": ["bis(2,4,6-trimethylbenzoyl)phenylphosphine oxide", "irganox 819", "irgamet 819", "omnirad 819"],
    "lap": ["lithium phenyl-2,4,6-trimethylbenzoylphosphinate", "lithium phenyl (2,4,6-trimethylbenzoyl)phosphinate"],
    "cq": ["camphorquinone", "bornane-2,3-dione"],
    "edab": ["ethyl 4-(dimethylamino)benzoate", "ethyl 4-dimethylaminobenzoate"],
    "mdea": ["n-methyldiethanolamine", "methyldiethanolamine"],
    "tempo": ["(2,2,6,6-tetramethylpiperidin-1-yl)oxyl", "tetramethylpiperidine 1-oxyl"],
    "photoinitiator": ["initiator", "pi", "photosensitizer", "radical initiator", "cleavage initiator"],
    "inhibitor": ["scavenger", "radical scavenger", "tempo", "hydroquinone", "mehq"],

    # Volumetric Additive Manufacturing & Techniques
    "vam": ["volumetric additive manufacturing", "computed axial lithography", "cal", "tomographic printing", "tomographic additive manufacturing", "tam"],
    "cal": ["computed axial lithography", "vam", "volumetric additive manufacturing", "tomographic printing"],
    "xolo": ["xolography", "dual-color", "dual-wavelength printing", "light sheet"],
    "dlp": ["digital light processing", "dmd", "digital micromirror device"],
    "sla": ["stereolithography", "vat photopolymerization"],

    # Optics & Parameters
    "wavelength": ["nm", "nanometer", "spectral line", "excitation wavelength"],
    "attenuation": ["absorbance", "optical density", "extinction", "transmittance", "penetration depth"],
    "viscosity": ["viscous", "rheology", "flowability", "centipoise", "pas"],
    "refractive index": ["optical matching", "refractive matching", "index matching", "ri"],
    "power": ["irradiance", "intensity", "mw/cm2", "flux", "dose", "exposure"],
}

# Reverse index for quick bidirectional lookup
_REVERSE_SCIENTIFIC_INDEX: Dict[str, Set[str]] = {}
for _canonical, _syns in SCIENTIFIC_SYNONYMS.items():
    _group = set([_canonical] + [s.lower() for s in _syns])
    for _term in _group:
        if _term not in _REVERSE_SCIENTIFIC_INDEX:
            _REVERSE_SCIENTIFIC_INDEX[_term] = set()
        _REVERSE_SCIENTIFIC_INDEX[_term].update(_group - {_term})


class ThesaurusManager:
    """Manages scientific and WordNet synonym expansion."""

    def __init__(self, enable_wordnet: bool = True):
        self.enable_wordnet = enable_wordnet
        self._wordnet = None
        self._wordnet_attempted = False

    def _get_wordnet(self):
        if not self.enable_wordnet:
            return None
        if not self._wordnet_attempted:
            self._wordnet_attempted = True
            try:
                from nltk.corpus import wordnet
                # Ensure wordnet data is available
                _ = wordnet.synsets("test")
                self._wordnet = wordnet
            except Exception as e:
                logger.warning(
                    "WordNet corpus is unavailable (%s); English synonym expansion will be disabled.",
                    e,
                )
                self._wordnet = None
        return self._wordnet

    def get_scientific_synonyms(self, term: str) -> List[str]:
        """Look up known scientific synonyms for a word or multi-word term."""
        if not term or not isinstance(term, str):
            return []
        term_clean = term.strip().lower()
        if not term_clean:
            return []
        if term_clean in _REVERSE_SCIENTIFIC_INDEX:
            return sorted(list(_REVERSE_SCIENTIFIC_INDEX[term_clean]))

        # Check if term is inside any canonical key or item
        matches = set()
        for key, syns in SCIENTIFIC_SYNONYMS.items():
            all_in_group = [key] + [s.lower() for s in syns]
            if any(term_clean == item or term_clean in item.split() for item in all_in_group):
                matches.update(item for item in all_in_group if item != term_clean)
        return sorted(list(matches))

    def get_wordnet_synonyms(self, term: str, max_count: int = 6) -> List[str]:
        if not term or not isinstance(term, str):
            return []
        if (
            max_count is None
            or isinstance(max_count, bool)
            or not isinstance(max_count, (int, float))
            or max_count <= 0
        ):
            return []
        wn = self._get_wordnet()
        if not wn:
            return []
        term_clean = term.strip().lower()
        if not term_clean:
            return []
        synonyms: Set[str] = set()
        try:
            for synset in wn.synsets(term_clean):
                for lemma in synset.lemmas():
                    name = lemma.name().replace("_", " ").lower()
                    if name != term_clean and not name.startswith(term_clean):
                        synonyms.add(name)
                        if len(synonyms) >= max_count:
                            return sorted(list(synonyms))
        except Exception as e:
            logger.warning("Failed to retrieve WordNet synonyms for '%s': %s", term_clean, e)
        return sorted(list(synonyms))

    def expand_query(self, query: str, use_wordnet: bool = True, use_scientific: bool = True) -> Dict[str, List[str]]:
        """
        Takes a search query and returns a mapping from each word/token to its synonyms.
        Also checks for multi-word scientific terms (e.g. 'refractive index', 'computed axial lithography').
        """
        results: Dict[str, List[str]] = {}
        if not query or not isinstance(query, str):
            return results
        query_clean = query.strip().lower()
        if not query_clean:
            return results

        # 1. Multi-word phrase check in scientific synonyms
        if use_scientific:
            sci_phrase_syns = self.get_scientific_synonyms(query_clean)
            if sci_phrase_syns:
                results[query_clean] = sci_phrase_syns

        # 2. Tokenize into individual alphanumeric words
        words = re.findall(r"\b[a-zA-Z0-9_\-]+\b", query_clean)
        for w in words:
            if len(w) < 3 or w in results:
                continue
            collected: Set[str] = set()
            if use_scientific:
                for s in self.get_scientific_synonyms(w):
                    collected.add(s)
            if use_wordnet:
                for s in self.get_wordnet_synonyms(w):
                    collected.add(s)
            if collected:
                results[w] = sorted(list(collected))

        return results
