import hashlib
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Set
from difflib import SequenceMatcher
from fuzzywuzzy import fuzz
import nltk
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
from config import DUPLICATE_THRESHOLD

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DuplicateDetector:
    """Handles duplicate detection using multiple similarity algorithms"""
    
    def __init__(self):
        # Download required NLTK data
        try:
            nltk.download('punkt', quiet=True)
            nltk.download('stopwords', quiet=True)
        except:
            pass
    
    def find_duplicates(self, documents: Dict[str, str]) -> Dict[str, Dict]:
        """
        Find duplicates among a collection of documents
        
        Args:
            documents: Dictionary mapping file_path -> extracted_text
            
        Returns:
            Dictionary mapping file_path -> duplicate_info
        """
        duplicate_info = {}
        file_paths = list(documents.keys())
        
        # Initialize all files as non-duplicates
        for file_path in file_paths:
            duplicate_info[file_path] = {
                "is_duplicate": False,
                "similarity_percentage": 0.0,
                "duplicate_reason": "No duplicates found",
                "master_file": file_path,
                "duplicates_of": []
            }
        
        # Compare each file with every other file
        for i, file1 in enumerate(file_paths):
            for j, file2 in enumerate(file_paths[i+1:], i+1):
                similarity_result = self._calculate_similarity(
                    documents[file1], documents[file2], file1, file2
                )
                
                if similarity_result["similarity_percentage"] > DUPLICATE_THRESHOLD * 100:
                    # Determine which is the master file
                    master_file = self._determine_master_file(file1, file2, documents)
                    duplicate_file = file2 if master_file == file1 else file1
                    
                    # Update duplicate information
                    duplicate_info[duplicate_file].update({
                        "is_duplicate": True,
                        "similarity_percentage": similarity_result["similarity_percentage"],
                        "duplicate_reason": similarity_result["reason"],
                        "master_file": master_file,
                        "duplicates_of": [master_file]
                    })
                    
                    # Add to master's duplicate list
                    if duplicate_file not in duplicate_info[master_file]["duplicates_of"]:
                        duplicate_info[master_file]["duplicates_of"].append(duplicate_file)
        
        return duplicate_info
    
    def _calculate_similarity(self, text1: str, text2: str, file1: str, file2: str) -> Dict:
        """Calculate similarity between two texts using multiple methods"""
        
        # Method 1: Exact hash comparison
        hash1 = hashlib.md5(text1.encode()).hexdigest()
        hash2 = hashlib.md5(text2.encode()).hexdigest()
        
        if hash1 == hash2:
            return {
                "similarity_percentage": 100.0,
                "reason": "Exact duplicate (identical hash)"
            }
        
        # Method 2: Sequence matcher (character-level)
        seq_similarity = SequenceMatcher(None, text1, text2).ratio() * 100
        
        # Method 3: Fuzzy string matching
        fuzzy_ratio = fuzz.ratio(text1, text2)
        fuzzy_partial = fuzz.partial_ratio(text1, text2)
        fuzzy_token_sort = fuzz.token_sort_ratio(text1, text2)
        fuzzy_token_set = fuzz.token_set_ratio(text1, text2)
        
        # Method 4: TF-IDF cosine similarity
        tfidf_similarity = self._calculate_tfidf_similarity(text1, text2)
        
        # Combine all similarity scores
        similarities = [seq_similarity, fuzzy_ratio, fuzzy_partial, 
                       fuzzy_token_sort, fuzzy_token_set, tfidf_similarity * 100]
        
        # Use weighted average (giving more weight to semantic similarity)
        weights = [0.15, 0.15, 0.15, 0.15, 0.15, 0.25]
        weighted_similarity = sum(s * w for s, w in zip(similarities, weights))
        
        # Determine reason based on similarity patterns
        reason = self._determine_duplicate_reason(similarities)
        
        return {
            "similarity_percentage": round(weighted_similarity, 2),
            "reason": reason
        }
    
    def _calculate_tfidf_similarity(self, text1: str, text2: str) -> float:
        """Calculate TF-IDF cosine similarity"""
        try:
            vectorizer = TfidfVectorizer(stop_words='english', max_features=1000)
            tfidf_matrix = vectorizer.fit_transform([text1, text2])
            similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
            return similarity
        except Exception:
            return 0.0
    
    def _determine_duplicate_reason(self, similarities: List[float]) -> str:
        """Determine the reason for duplication based on similarity patterns"""
        seq_sim, fuzzy_ratio, fuzzy_partial, fuzzy_token_sort, fuzzy_token_set, tfidf_sim = similarities
        
        if seq_sim > 95:
            return "Nearly identical content"
        elif fuzzy_ratio > 90:
            return "Minor formatting differences"
        elif fuzzy_token_sort > 85:
            return "Same content, different word order"
        elif fuzzy_token_set > 85:
            return "Similar content with additions/deletions"
        elif tfidf_sim > 85:
            return "Semantically similar content"
        else:
            return "Low similarity - likely not duplicates"
    
    def _determine_master_file(self, file1: str, file2: str, documents: Dict[str, str]) -> str:
        """Determine which file is the master based on various criteria"""
        path1, path2 = Path(file1), Path(file2)
        text1, text2 = documents[file1], documents[file2]
        
        # Criteria for determining master file:
        # 1. File with more content (longer text)
        if len(text1) != len(text2):
            return file1 if len(text1) > len(text2) else file2
        
        # 2. File with earlier creation/modification time
        try:
            stat1, stat2 = path1.stat(), path2.stat()
            if stat1.st_mtime != stat2.st_mtime:
                return file1 if stat1.st_mtime < stat2.st_mtime else file2
        except:
            pass
        
        # 3. File with simpler name (shorter path, fewer special characters)
        if len(file1) != len(file2):
            return file1 if len(file1) < len(file2) else file2
        
        # 4. Alphabetical order as final fallback
        return file1 if file1 < file2 else file2
    
    def calculate_content_hash(self, text: str) -> str:
        """Calculate content hash for quick duplicate detection"""
        # Normalize text for better hash comparison
        normalized_text = ' '.join(text.lower().split())
        return hashlib.sha256(normalized_text.encode()).hexdigest()
    
    def group_by_hash(self, documents: Dict[str, str]) -> Dict[str, List[str]]:
        """Group documents by content hash for quick duplicate identification"""
        hash_groups = {}
        
        for file_path, text in documents.items():
            content_hash = self.calculate_content_hash(text)
            if content_hash not in hash_groups:
                hash_groups[content_hash] = []
            hash_groups[content_hash].append(file_path)
        
        # Return only groups with multiple files (potential duplicates)
        return {h: files for h, files in hash_groups.items() if len(files) > 1}