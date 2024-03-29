import concurrent.futures
import os
import re
from tkinter import filedialog, StringVar, Text, messagebox
from typing import List, Generator, Optional, Tuple, Union

import PyPDF2
import customtkinter as ctk
import docx
from nltk.corpus import wordnet, words
from nltk.tokenize import sent_tokenize


class DocumentProcessor:
    def __init__(self, search_folder: str, search_word: str, restrict: bool, tokenized_sentences=None):
        """
        Initialize the DocumentProcessor.

        :param search_folder: The folder path to search for documents.
        :param search_word: The word to search for.
        :param restrict: Whether to restrict the search or not.
        :param tokenized_sentences: A pre-tokenized dictionary of sentences for each file (default: None).
        """
        self.search_folder = search_folder
        self._search_word = search_word
        self._restrict_search = restrict
        self.synonyms = self.get_synonyms() if not self._restrict_search else []
        self.docs_count = 0
        self.all_docs = 0
        self.tokenized_sentences = tokenized_sentences if tokenized_sentences is not None else {}
        self.get_all_docs()

    @property
    def restrict_search(self):
        return self._restrict_search

    @restrict_search.setter
    def restrict_search(self, new_restrict_search_bool: bool):
        self._restrict_search = new_restrict_search_bool
        self.synonyms = self.get_synonyms() if not self._restrict_search else []

    @property
    def search_word(self):
        return self._search_word

    @search_word.setter
    def search_word(self, new_search_word: str):
        self._search_word = new_search_word
        self.synonyms = self.get_synonyms() if not self._restrict_search else []

    def get_synonyms(self) -> List[str]:
        """
        Get the synonyms of the search word using WordNet.

        :return: A list of synonyms.
        """
        synonyms = set()
        for syn in wordnet.synsets(self._search_word):
            for lemma in syn.lemmas():
                synonyms.add(lemma.name())
        return list(synonyms)

    def get_all_docs(self) -> None:
        """
        Count the total number of documents in the search folder.
        """
        for root_dir, _, files in os.walk(self.search_folder):
            for file in files:
                if file.endswith(".pdf") or file.endswith(".docx"):
                    self.all_docs += 1

    def get_files(self) -> Generator[str, None, None]:
        """
        Get the file paths of the documents to search.

        :yield: The file path of each document.
        """
        for root_dir, _, files in os.walk(self.search_folder):
            for file in files:
                if file.endswith(".pdf") or file.endswith(".docx"):
                    # self.docs_count += 1
                    # yield os.path.join(root_dir, file)
                    file_path = os.path.join(root_dir, file)
                    normalized_path = os.path.normpath(file_path)
                    self.docs_count += 1
                    yield normalized_path

    def search_file(self, file_path) -> Optional[List[Tuple[str, int, List[Tuple[int, int, str]]]]]:
        """
        Search a specific file for occurrences of the search word.

        :param file_path: The path of the file to search.
        :return: A list containing information about the matches found in the file.
                 Each item in the list is a tuple with the file path, the number of matches, and a list of matches.
                 Each match is represented by a tuple with the page number, sentence number, and the highlighted line.
                 Returns None if no matches are found.
        """
        if file_path not in self.tokenized_sentences:

            with open(file_path, "rb") as file:
                if file_path.endswith(".pdf"):
                    reader = PyPDF2.PdfReader(file)
                    page_texts = [(page_num, page.extract_text()) for page_num, page in enumerate(reader.pages)]
                elif file_path.endswith(".docx"):
                    doc = docx.Document(file_path)
                    page_texts = [(page_num, paragraph.text) for page_num, paragraph in enumerate(doc.paragraphs)]
                else:
                    return None

                lines = []
                for page_num, text in page_texts:
                    for sentence in sent_tokenize(text):
                        lines.append((page_num, sentence))
                self.tokenized_sentences[file_path] = lines
        else:
            lines = self.tokenized_sentences[file_path]

        if self._restrict_search:
            matches = self.process_text(lines, restrict=True)
        else:
            matches = self.process_text(lines)
        return [[file_path, len(matches), matches]] if matches else None

    def process_text(self, lines: List[Tuple[int, str]], restrict: bool = False) -> list[list[Union[int, str]]]:
        """
        Process a list of text and search for occurrences of the search word.

        :param lines: The list of text to process, containing tuples with page numbers and sentences.
        :param restrict: Whether to restrict the search to only direct matches.
        :return: A list of matches found in the text.
                 Each match is represented by a tuple with the page number, sentence number, and the highlighted line.
        """
        matches = []
        prev_page_num = 0
        sentence_num_in_page = 0

        for line_num, (page_num, line) in enumerate(lines):
            if prev_page_num != page_num:
                sentence_num_in_page = 0
                prev_page_num = page_num

            if restrict:
                search_pattern = r'\b{}\b'.format(self._search_word)
            else:
                search_pattern = r'({}|{})'.format(self._search_word, '|'.join(map(re.escape, self.synonyms)))
            found_words = re.findall(search_pattern, line, re.IGNORECASE)

            if found_words:
                line_highlighted = re.sub(search_pattern, lambda m: f'<<{m.group(0)}>>', line)
                matches.append([page_num, sentence_num_in_page, line_highlighted])

            sentence_num_in_page += 1

        return matches

    def process_files(self) -> List[Tuple[str, int, List[Tuple[int, int, str]]]]:
        """
        Search all the files in the search folder for occurrences of the search word.

        :return: A list containing information about the matches found in each file.
        """
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future_to_file = {executor.submit(self.search_file, file_path): file_path for file_path in self.get_files()}
            results = []
            for future in concurrent.futures.as_completed(future_to_file):
                result = future.result()
                if result:
                    results.extend(result)
            return results


class LWSgui(ctk.CTk):
    def __init__(self, *args, **kwargs):

        """
        Initialize the LWSgui.

        """
        super().__init__(*args, **kwargs)
        self.tokenized_sentences = None
        self.current_directory = None
        self.processor = None

        self.minsize(800, 250)

        if ctk.get_appearance_mode() == 'Dark':
            self.iconbitmap('Media/white.ico')
        if ctk.get_appearance_mode() == 'Light':
            self.iconbitmap('Media/dark.ico')
        self.title("LWS - LuminousWordSearch")

        self.left_frame = LeftFrame(self)
        self.left_frame.grid(row=0, column=0, sticky="nsew")

        self.right_frame = RightFrame(self)
        self.right_frame.grid(row=0, column=1, sticky="nsew")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.mainloop()

    def display_result(self, text: str, color=None) -> None:
        """
        Display the result in the result text box.

        :param text: The text to display.
        :param color: The color of the text (optional).
        :return: None
        """
        if color:
            self.right_frame.result_text.tag_configure(color, foreground=color)
            self.right_frame.result_text.insert("end", text, color)
        else:
            self.right_frame.result_text.insert("end", text)
        self.right_frame.result_text.update_idletasks()

    def browse_folder(self) -> None:
        """
        Open a folder selection dialog and set the selected folder as the search folder.
        :return: None
        """
        folder = filedialog.askdirectory()
        if folder:
            self.left_frame.search_folder.set(folder)

    def search(self, _event=None) -> None:
        """
        Perform the search based on the provided search word and folder.
        :param _event: Tkinter event (default: None)
        :return: None
        """
        search_word = self.left_frame.word_entry.get()
        folder = self.left_frame.search_folder.get()
        restrict_search = self.left_frame.restrict_var.get()

        if not search_word or not folder:
            messagebox.showwarning("Warning", "Please provide a search word and folder.")
            return

        if not search_word.lower() in words.words():
            response = messagebox.askquestion("Word Not Found",
                                              "The search word is not spelled correctly. Do you want to continue?")
            if response == "no":
                return

        # Check if the processor should be updated
        if self.processor is None or self.current_directory != folder:
            self.processor = DocumentProcessor(folder, search_word, restrict_search)
            self.current_directory = folder

        if search_word != self.processor.search_word:
            self.processor.search_word = search_word

        if restrict_search != self.processor.restrict_search:
            self.processor.restrict_search = restrict_search

        self.left_frame.synonyms_frame.synonyms_text_box.delete('1.0', 'end')
        self.right_frame.result_text.delete('1.0', 'end')
        self.display_result("Results of searching for ", "black")
        self.display_result(search_word, "orange")
        self.display_result("\n\n", "black")
        self.left_frame.progress_bar.set(0)
        self.update()

        synonyms = self.processor.synonyms
        self.left_frame.synonyms_frame.synonyms_text_box.tag_configure("purple", foreground="purple")
        if synonyms:
            for tic, synonym in enumerate(synonyms):
                self.left_frame.synonyms_frame.synonyms_text_box.insert("end", synonym, "purple")
                if tic < len(synonyms) - 1:
                    self.left_frame.synonyms_frame.synonyms_text_box.insert("end", ', ', "purple")
        if not synonyms:
            if not restrict_search:
                self.left_frame.synonyms_frame.synonyms_text_box.tag_configure("red", foreground="red")
                self.left_frame.synonyms_frame.synonyms_text_box.insert("end",
                                                                        'You are in the non-optimized and dangerous'
                                                                        ' mode!\nAnything you find can and will be '
                                                                        'used against you in a court of science!',
                                                                        "red")
            if restrict_search:
                self.left_frame.synonyms_frame.synonyms_text_box.insert("end",
                                                                        'Restricted mode, no synonyms are allowed'
                                                                        '\nYou are amazing as always, love you!',
                                                                        "purple")
        self.left_frame.synonyms_frame.synonyms_text_box.update_idletasks()
        self.left_frame.progress_bar.set(0)
        total_matches = 0
        results = self.processor.process_files()
        for result in results:
            file_path, _, line_lists = result
            self.left_frame.progress_bar.set(self.processor.docs_count / self.processor.all_docs)
            self.update_idletasks()
            if not result:
                pass
            file_name = os.path.basename(file_path)
            self.display_result(f"\n{file_name}:\n", "blue")

            # Count matches and update the total
            direct_matches, synonym_matches = self.count_matches([result], search_word)
            total_matches += direct_matches + synonym_matches

            # Display match count
            self.display_result(f"Direct matches: ")
            self.display_result(f"{direct_matches}\n", "red")
            self.display_result(f"Synonym matches: ")
            self.display_result(f"{synonym_matches}\n", "purple")

            for line_list in line_lists:
                self.display_matches(line_list, search_word, synonyms)

    def display_matches(self, line_list: Tuple[int, int, str], search_word: str, synonyms: List[str]) -> None:
        """
        Display matches in the result_text widget.

        :param line_list: Tuple containing page number, line number, and highlighted sentence
        :param search_word: The word being searched for
        :param synonyms: List of synonyms for the search_word
        :return: None
        """
        page_num, line_num, line_highlighted = line_list
        self.display_result(f"Page {page_num + 1} Sentence {line_num + 1}: ", 'green')
        all_words = line_highlighted.split('<<')
        for word in all_words:
            if '>>' in word:
                word, remaining = word.split('>>', 1)
                if word.lower() == search_word.lower():
                    self.display_result(f"{word}", "red")
                elif word.lower() in [synonym.lower() for synonym in synonyms]:
                    self.display_result(f"{word}", "magenta")
                else:
                    self.display_result(word)
                self.display_result(remaining)
            else:
                self.display_result(word)
        self.display_result("\n")

    @staticmethod
    def count_matches(result: List[Tuple[str, int, List[Tuple[int, int, str]]]], search_word: str) -> Tuple[int, int]:
        """
        Count direct matches and synonym matches in the result.

        :param result: List of results from searching the file
        :param search_word: The word being searched for
        :return: Tuple containing direct_matches and synonym_matches
        """
        direct_matches = 0
        synonym_matches = 0
        for r in result:
            _, _, line_lists = r
            for line_list in line_lists:
                _, _, line_highlighted = line_list
                if search_word.lower() in line_highlighted.lower():
                    direct_matches += 1
                else:
                    synonym_matches += 1
        return direct_matches, synonym_matches

    def update_font(self, action, value=None):
        """
        Update the font size or font face of the result_text and synonyms_text_box based on button clicks or option
        menu selection.

        :param action: The action to perform, either "increase", "decrease", or "change_face".
        :param value: The font face to change to when the action is "change_face".
        :return: None
        """
        widgets = [self.right_frame.result_text, self.left_frame.synonyms_frame.synonyms_text_box]

        for widget in widgets:
            current_font = widget.cget("font")
            current_font_size = widget.tk.call("font", "actual", current_font, "-size")
            current_font_face = widget.tk.call("font", "actual", current_font, "-family")
            new_font_size = current_font_size

            if action == "increase":
                new_font_size += 1
            elif action == "decrease":
                new_font_size -= 1
            elif action == "change_face":
                current_font_face = value

            if new_font_size > 0:
                widget.configure(font=(current_font_face, new_font_size))


class LeftFrame(ctk.CTkFrame):
    """
    The left frame of the LWSgui containing the search controls.
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(master=parent, *args, **kwargs)
        self.parent = parent
        self.search_folder = StringVar()
        self.folder_label = ctk.CTkLabel(self, text="Search folder:")
        self.folder_entry = ctk.CTkEntry(self, textvariable=self.search_folder)
        self.browse_button = ctk.CTkButton(self, text="Browse", command=self.parent.browse_folder)
        self.word_label = ctk.CTkLabel(self, text="Search word:")
        self.word_entry = ctk.CTkEntry(self)
        self.word_entry.bind("<KeyPress>", self.no_space_keypress)
        self.word_entry.bind("<Return>", self.parent.search)
        self.restrict_var = ctk.BooleanVar()
        self.restrict_check = ctk.CTkCheckBox(self, text="Restrict search", variable=self.restrict_var)
        self.search_button = ctk.CTkButton(self, text="Search", command=self.parent.search)
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.set(0)
        self.folder_label.grid(row=0, column=0, padx=5, pady=5)
        self.folder_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)
        self.word_label.grid(row=1, column=0, padx=5, pady=5)
        self.word_entry.grid(row=1, column=1, padx=5, pady=5)
        self.restrict_check.grid(row=1, column=2, padx=5, pady=5)
        self.search_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)
        self.progress_bar.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        self.font_controls_frame = FontControlFrame(self)
        self.font_controls_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

        self.synonyms_frame = SynonymsFrame(self)
        self.synonyms_frame.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    @staticmethod
    def no_space_keypress(event):
        if event.keysym == "space":
            return "break"


class RightFrame(ctk.CTkFrame):
    """
    Right frame of the LWSgui containing the result text box.
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(master=parent, *args, **kwargs)
        self.parent = parent
        self.result_text = Text(self, wrap='word')
        self.result_text.grid(row=0, column=0, pady=5, sticky="nsew")
        self.result_text.tag_configure("orange", foreground="orange")
        self.scroll_bar = ctk.CTkScrollbar(self, command=self.result_text.yview)
        self.scroll_bar.grid(row=0, column=1, sticky='nsew')
        self.result_text['yscrollcommand'] = self.scroll_bar.set
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)


class FontControlFrame(ctk.CTkFrame):
    """
    Frame containing font size and font face controls for the result_text and synonyms_text_box widgets.
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(master=parent, *args, **kwargs)
        self.parent = parent
        self.font_size_label = ctk.CTkLabel(self, text="Font Size:")
        self.font_size_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        self.increase_font_button = ctk.CTkButton(self, text="+", width=20, height=20,
                                                  command=lambda: self.parent.parent.update_font("increase"))
        self.decrease_font_button = ctk.CTkButton(self, text="-", width=20, height=20,
                                                  command=lambda: self.parent.parent.update_font("decrease"))
        self.increase_font_button.grid(row=1, column=0, padx=(0, 2), pady=5)
        self.decrease_font_button.grid(row=1, column=1, padx=(2, 0), pady=5)

        # Font face OptionMenu
        self.font_face_label = ctk.CTkLabel(self, text="Font:")
        self.font_face_label.grid(row=0, column=3, padx=5, pady=5, sticky="EW")
        self.font_var = ctk.StringVar()
        self.available_fonts = ["Arial", "Courier", "Helvetica", "Times New Roman", "Verdana", "Comic Sans MS",
                                "Georgia", "Lucida Sans", "Tahoma", "Trebuchet MS"]
        self.font_var.set(self.available_fonts[0])

        def option_menu_callback(choice):
            self.parent.parent.update_font("change_face", choice)

        self.font_option_menu = ctk.CTkOptionMenu(master=self,
                                                  values=self.available_fonts,
                                                  command=option_menu_callback,
                                                  variable=self.font_var)
        self.font_option_menu.grid(row=1, column=3, columnspan=2, padx=5, pady=5, sticky="EW")
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=5)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=9)


class SynonymsFrame(ctk.CTkFrame):
    """
    Frame containing a text box to display the synonyms of the search word.
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(master=parent, *args, **kwargs)
        self.parent = parent
        self.configure(width=350, height=250)
        self.pack_propagate(False)
        self.synonyms_text_box_label = ctk.CTkLabel(self, text="Synonyms of searching word")
        self.synonyms_text_box = Text(self, wrap='word', height=1, width=1)
        self.synonyms_text_box_label.pack(pady=5)
        self.synonyms_text_box.pack(side='left', pady=0, fill='both', expand=True)
        self.synonyms_text_box_scroll_bar = ctk.CTkScrollbar(self, command=self.synonyms_text_box.yview)
        self.synonyms_text_box_scroll_bar.pack(side='right', fill='both')
        self.synonyms_text_box['yscrollcommand'] = self.synonyms_text_box_scroll_bar.set


if __name__ == "__main__":
    LWSgui()
