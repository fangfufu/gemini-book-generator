import sys
import yaml
import logging
import threading
from PySide6.QtCore import QObject, Signal, Slot, QEvent
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QTextEdit, QLineEdit, QLabel,
    QCheckBox, QSpinBox, QDoubleSpinBox, QComboBox, QFormLayout,
    QGroupBox, QStackedWidget, QFileDialog
)

from book_generator.config import load_config, get_config_path
from book_generator.orchestrator import run_generation_process

# Custom logging handler to emit signals
class QTextEditHandler(QObject, logging.Handler):
    new_log_record = Signal(str)

    def __init__(self):
        super().__init__()
        QObject.__init__(self)
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        msg = self.format(record)
        self.new_log_record.emit(msg)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Book Generator Configuration")
        self.setGeometry(100, 100, 800, 700)

        self.config_path = get_config_path()
        self.config = load_config()
        self.widgets = {}

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        self._create_generation_tab()
        self._create_api_tab()
        self._create_style_tab()
        self._create_output_debug_tab()

        self._create_log_area()
        self._create_buttons()

        self.load_config_to_gui()

    def _create_generation_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        self.tabs.addTab(tab, "Generation")

        self.widgets['generation_params'] = {}

        # main_topic
        self.widgets['generation_params']['main_topic'] = QTextEdit()
        layout.addRow("Main Topic:", self.widgets['generation_params']['main_topic'])

        # author_name, author_gender
        author_layout = QHBoxLayout()
        self.widgets['generation_params']['author_name'] = QLineEdit()
        self.widgets['generation_params']['author_gender'] = QComboBox()
        self.widgets['generation_params']['author_gender'].addItems(["", "male", "female", "other"])
        author_layout.addWidget(self.widgets['generation_params']['author_name'])
        author_layout.addWidget(QLabel("Gender:"))
        author_layout.addWidget(self.widgets['generation_params']['author_gender'])
        layout.addRow("Author:", author_layout)

        # is_fiction, fiction_chapter_count
        fiction_layout = QHBoxLayout()
        self.widgets['generation_params']['is_fiction'] = QCheckBox("Is Fiction")
        self.widgets['generation_params']['fiction_chapter_count'] = QSpinBox()
        self.widgets['generation_params']['fiction_chapter_count'].setRange(1, 200)
        fiction_layout.addWidget(self.widgets['generation_params']['is_fiction'])
        fiction_layout.addWidget(QLabel("Chapters:"))
        fiction_layout.addWidget(self.widgets['generation_params']['fiction_chapter_count'])
        layout.addRow(fiction_layout)

        # Other fields
        self.widgets['generation_params']['book_title'] = QLineEdit()
        layout.addRow("Book Title:", self.widgets['generation_params']['book_title'])
        self.widgets['generation_params']['book_subtitle'] = QLineEdit()
        layout.addRow("Book Subtitle:", self.widgets['generation_params']['book_subtitle'])
        self.widgets['generation_params']['setting'] = QTextEdit()
        layout.addRow("Setting:", self.widgets['generation_params']['setting'])
        self.widgets['generation_params']['writing_tone'] = QLineEdit()
        layout.addRow("Writing Tone:", self.widgets['generation_params']['writing_tone'])
        self.widgets['generation_params']['random_topic_seed'] = QLineEdit()
        layout.addRow("Random Topic Seed:", self.widgets['generation_params']['random_topic_seed'])
        self.widgets['generation_params']['length_modifier'] = QLineEdit()
        layout.addRow("Length Modifier:", self.widgets['generation_params']['length_modifier'])

        # Booleans
        self.widgets['generation_params']['generate_character_list'] = QCheckBox("Generate Character List")
        layout.addRow(self.widgets['generation_params']['generate_character_list'])
        self.widgets['generation_params']['generate_foreword'] = QCheckBox("Generate Foreword")
        layout.addRow(self.widgets['generation_params']['generate_foreword'])
        self.widgets['generation_params']['generate_acknowledgements'] = QCheckBox("Generate Acknowledgements")
        layout.addRow(self.widgets['generation_params']['generate_acknowledgements'])
        self.widgets['generation_params']['generate_appendix'] = QCheckBox("Generate Appendix")
        layout.addRow(self.widgets['generation_params']['generate_appendix'])
        self.widgets['generation_params']['include_about_author_in_main_book'] = QCheckBox("Include 'About Author' in book")
        layout.addRow(self.widgets['generation_params']['include_about_author_in_main_book'])

    def _create_api_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.tabs.addTab(tab, "API")

        self.widgets['api_settings'] = {'gemini': {}, 'ollama': {}}

        # Provider selection
        provider_layout = QFormLayout()
        self.widgets['api_settings']['provider'] = QComboBox()
        self.widgets['api_settings']['provider'].addItems(["gemini", "ollama"])
        provider_layout.addRow("API Provider:", self.widgets['api_settings']['provider'])
        layout.addLayout(provider_layout)

        # Provider-specific settings
        self.provider_stack = QStackedWidget()
        gemini_widget = self._create_gemini_group()
        ollama_widget = self._create_ollama_group()
        self.provider_stack.addWidget(gemini_widget)
        self.provider_stack.addWidget(ollama_widget)
        layout.addWidget(self.provider_stack)
        self.widgets['api_settings']['provider'].currentIndexChanged.connect(self.provider_stack.setCurrentIndex)

        # Common API settings
        common_group = QGroupBox("Common API Settings")
        common_layout = QFormLayout(common_group)
        self.widgets['api_settings']['base_cache_dir'] = QLineEdit()
        common_layout.addRow("Cache Directory:", self.widgets['api_settings']['base_cache_dir'])
        self.widgets['api_settings']['default_max_retries'] = QSpinBox()
        common_layout.addRow("Default Max Retries:", self.widgets['api_settings']['default_max_retries'])
        self.widgets['api_settings']['default_retry_delay_seconds'] = QSpinBox()
        common_layout.addRow("Default Retry Delay (s):", self.widgets['api_settings']['default_retry_delay_seconds'])
        layout.addWidget(common_group)

    def _create_gemini_group(self):
        group = QGroupBox("Gemini Settings")
        layout = QFormLayout(group)
        self.widgets['api_settings']['gemini']['model'] = QLineEdit()
        layout.addRow("Model:", self.widgets['api_settings']['gemini']['model'])
        self.widgets['api_settings']['gemini']['temperature'] = QDoubleSpinBox()
        layout.addRow("Temperature:", self.widgets['api_settings']['gemini']['temperature'])
        self.widgets['api_settings']['gemini']['max_retries'] = QSpinBox()
        layout.addRow("Max Retries:", self.widgets['api_settings']['gemini']['max_retries'])
        self.widgets['api_settings']['gemini']['retry_delay_seconds'] = QSpinBox()
        layout.addRow("Retry Delay (s):", self.widgets['api_settings']['gemini']['retry_delay_seconds'])
        return group

    def _create_ollama_group(self):
        group = QGroupBox("Ollama Settings")
        layout = QFormLayout(group)
        self.widgets['api_settings']['ollama']['base_url'] = QLineEdit()
        layout.addRow("Base URL:", self.widgets['api_settings']['ollama']['base_url'])
        self.widgets['api_settings']['ollama']['model'] = QLineEdit()
        layout.addRow("Model:", self.widgets['api_settings']['ollama']['model'])
        self.widgets['api_settings']['ollama']['tokenizer_model'] = QLineEdit()
        layout.addRow("Tokenizer Model:", self.widgets['api_settings']['ollama']['tokenizer_model'])
        self.widgets['api_settings']['ollama']['temperature'] = QDoubleSpinBox()
        layout.addRow("Temperature:", self.widgets['api_settings']['ollama']['temperature'])
        self.widgets['api_settings']['ollama']['context_window_size'] = QSpinBox()
        self.widgets['api_settings']['ollama']['context_window_size'].setMaximum(99999)
        layout.addRow("Context Window:", self.widgets['api_settings']['ollama']['context_window_size'])
        return group

    def _create_style_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        self.tabs.addTab(tab, "Style")
        self.widgets['style_params'] = {'margins_mm': {}}

        self.widgets['style_params']['font_name'] = QLineEdit()
        layout.addRow("Font Name:", self.widgets['style_params']['font_name'])
        self.widgets['style_params']['font_size'] = QDoubleSpinBox()
        layout.addRow("Font Size (pt):", self.widgets['style_params']['font_size'])
        self.widgets['style_params']['page_size_preset'] = QComboBox()
        self.widgets['style_params']['page_size_preset'].addItems(["6x9", "A4"])
        layout.addRow("Page Size:", self.widgets['style_params']['page_size_preset'])

        margins_group = QGroupBox("Margins (mm)")
        margins_layout = QFormLayout(margins_group)
        self.widgets['style_params']['margins_mm']['top'] = QSpinBox()
        margins_layout.addRow("Top:", self.widgets['style_params']['margins_mm']['top'])
        self.widgets['style_params']['margins_mm']['bottom'] = QSpinBox()
        margins_layout.addRow("Bottom:", self.widgets['style_params']['margins_mm']['bottom'])
        self.widgets['style_params']['margins_mm']['left'] = QSpinBox()
        margins_layout.addRow("Left (Outside):", self.widgets['style_params']['margins_mm']['left'])
        self.widgets['style_params']['margins_mm']['right'] = QSpinBox()
        margins_layout.addRow("Right (Inside):", self.widgets['style_params']['margins_mm']['right'])
        self.widgets['style_params']['margins_mm']['gutter'] = QSpinBox()
        margins_layout.addRow("Gutter:", self.widgets['style_params']['margins_mm']['gutter'])
        layout.addRow(margins_group)

        math_group = QGroupBox("Math Rendering")
        math_layout = QFormLayout(math_group)
        self.widgets['style_params']['inline_math_height_multiplier'] = QDoubleSpinBox()
        math_layout.addRow("Inline Math Height Multiplier:", self.widgets['style_params']['inline_math_height_multiplier'])
        self.widgets['style_params']['display_math_height_inches'] = QDoubleSpinBox()
        math_layout.addRow("Display Math Height (in):", self.widgets['style_params']['display_math_height_inches'])
        self.widgets['style_params']['inline_math_vertical_offset_half_points'] = QSpinBox()
        self.widgets['style_params']['inline_math_vertical_offset_half_points'].setRange(-100, 100)
        math_layout.addRow("Inline Math Vertical Offset (half-points):", self.widgets['style_params']['inline_math_vertical_offset_half_points'])
        layout.addRow(math_group)


    def _create_output_debug_tab(self):
        tab = QWidget()
        layout = QFormLayout(tab)
        self.tabs.addTab(tab, "Output & Debug")
        self.widgets['debug_options'] = {}

        self.widgets['output_dir'] = QLineEdit()
        layout.addRow("Output Directory:", self.widgets['output_dir'])

        self.widgets['debug_options']['verbose_debug'] = QCheckBox("Enable Verbose Debug Logging")
        layout.addRow(self.widgets['debug_options']['verbose_debug'])
        self.widgets['debug_options']['save_intermediate_html'] = QCheckBox("Save Intermediate HTML")
        layout.addRow(self.widgets['debug_options']['save_intermediate_html'])
        self.widgets['debug_options']['log_output_path'] = QLineEdit()
        layout.addRow("Log File Path:", self.widgets['debug_options']['log_output_path'])

    def _create_log_area(self):
        log_group = QGroupBox("Logs")
        log_layout = QVBoxLayout(log_group)
        self.log_text_area = QTextEdit()
        self.log_text_area.setReadOnly(True)
        log_layout.addWidget(self.log_text_area)
        self.main_layout.addWidget(log_group)

        # Setup logging handler
        self.log_handler = QTextEditHandler()
        self.log_handler.new_log_record.connect(self.append_log_message)
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)


    def _create_buttons(self):
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Save Config")
        self.save_button.clicked.connect(self.save_config_from_gui)
        self.run_button = QPushButton("Run Generation")
        self.run_button.clicked.connect(self.run_generation)
        button_layout.addStretch()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.run_button)
        self.main_layout.addLayout(button_layout)

    def load_config_to_gui(self):
        # Generation Params
        gen_params = self.config.get('generation_params', {})
        for key, widget in self.widgets['generation_params'].items():
            value = gen_params.get(key)
            if value is not None:
                if isinstance(widget, QLineEdit) or isinstance(widget, QTextEdit):
                    widget.setText(str(value))
                elif isinstance(widget, QCheckBox):
                    widget.setChecked(bool(value))
                elif isinstance(widget, QSpinBox) or isinstance(widget, QDoubleSpinBox):
                    widget.setValue(value)
                elif isinstance(widget, QComboBox):
                    widget.setCurrentText(str(value))

        # API Settings
        api_settings = self.config.get('api_settings', {})
        self.widgets['api_settings']['provider'].setCurrentText(api_settings.get('provider', 'gemini'))
        for provider in ['gemini', 'ollama']:
            provider_settings = api_settings.get(provider, {})
            for key, widget in self.widgets['api_settings'][provider].items():
                value = provider_settings.get(key)
                if value is not None:
                    if isinstance(widget, QLineEdit): widget.setText(str(value))
                    else: widget.setValue(value)

        self.widgets['api_settings']['base_cache_dir'].setText(api_settings.get('base_cache_dir', 'api_cache'))
        self.widgets['api_settings']['default_max_retries'].setValue(api_settings.get('default_max_retries', 10))
        self.widgets['api_settings']['default_retry_delay_seconds'].setValue(api_settings.get('default_retry_delay_seconds', 30))

        # Style Params
        style_params = self.config.get('style_params', {})
        for key, widget in self.widgets['style_params'].items():
            if key == 'margins_mm': continue
            value = style_params.get(key)
            if value is not None:
                if isinstance(widget, QLineEdit): widget.setText(str(value))
                elif isinstance(widget, QComboBox): widget.setCurrentText(str(value))
                else: widget.setValue(value)

        margins = style_params.get('margins_mm', {})
        for key, widget in self.widgets['style_params']['margins_mm'].items():
            widget.setValue(margins.get(key, 15))

        # Output & Debug
        self.widgets['output_dir'].setText(self.config.get('output_dir', 'output'))
        debug_options = self.config.get('debug_options', {})
        self.widgets['debug_options']['verbose_debug'].setChecked(debug_options.get('verbose_debug', False))
        self.widgets['debug_options']['save_intermediate_html'].setChecked(debug_options.get('save_intermediate_html', False))
        self.widgets['debug_options']['log_output_path'].setText(debug_options.get('log_output_path', ''))

    def save_config_from_gui(self):
        # Generation Params
        for key, widget in self.widgets['generation_params'].items():
            if isinstance(widget, QLineEdit) or isinstance(widget, QTextEdit):
                self.config['generation_params'][key] = widget.toPlainText() if isinstance(widget, QTextEdit) else widget.text()
            elif isinstance(widget, QCheckBox):
                self.config['generation_params'][key] = widget.isChecked()
            elif isinstance(widget, QSpinBox) or isinstance(widget, QDoubleSpinBox):
                self.config['generation_params'][key] = widget.value()
            elif isinstance(widget, QComboBox):
                self.config['generation_params'][key] = widget.currentText()

        # API Settings
        api_settings = self.config['api_settings']
        api_settings['provider'] = self.widgets['api_settings']['provider'].currentText()
        for provider in ['gemini', 'ollama']:
            for key, widget in self.widgets['api_settings'][provider].items():
                api_settings[provider][key] = widget.text() if isinstance(widget, QLineEdit) else widget.value()

        api_settings['base_cache_dir'] = self.widgets['api_settings']['base_cache_dir'].text()
        api_settings['default_max_retries'] = self.widgets['api_settings']['default_max_retries'].value()
        api_settings['default_retry_delay_seconds'] = self.widgets['api_settings']['default_retry_delay_seconds'].value()

        # Style Params
        style_params = self.config['style_params']
        for key, widget in self.widgets['style_params'].items():
            if key == 'margins_mm': continue
            style_params[key] = widget.text() if isinstance(widget, QLineEdit) else widget.currentText() if isinstance(widget, QComboBox) else widget.value()

        for key, widget in self.widgets['style_params']['margins_mm'].items():
            style_params['margins_mm'][key] = widget.value()

        # Output & Debug
        self.config['output_dir'] = self.widgets['output_dir'].text()
        debug_options = self.config['debug_options']
        debug_options['verbose_debug'] = self.widgets['debug_options']['verbose_debug'].isChecked()
        debug_options['save_intermediate_html'] = self.widgets['debug_options']['save_intermediate_html'].isChecked()
        debug_options['log_output_path'] = self.widgets['debug_options']['log_output_path'].text()

        try:
            with open(self.config_path, 'w') as f:
                yaml.dump(self.config, f, default_flow_style=False, sort_keys=False)
            self.log_text_area.append("Configuration saved successfully.")
        except Exception as e:
            self.log_text_area.append(f"Error saving configuration: {e}")

    def run_generation(self):
        self.save_config_from_gui()
        self.log_text_area.append("Starting book generation process...")
        self.run_button.setEnabled(False)
        self.save_button.setEnabled(False)

        # Run in a separate thread
        thread = threading.Thread(target=self._run_process_in_thread)
        thread.start()

    def _run_process_in_thread(self):
        try:
            run_generation_process()
            self.log_handler.new_log_record.emit("Generation process completed successfully.")
        except Exception as e:
            self.log_handler.new_log_record.emit(f"An error occurred during generation: {e}")
        finally:
            # Re-enable buttons on the main thread
            QApplication.instance().postEvent(self, QEvent(QEvent.Type.User + 1))

    def event(self, event):
        if event.type() == QEvent.Type.User + 1:
            self.run_button.setEnabled(True)
            self.save_button.setEnabled(True)
            return True
        return super().event(event)

    @Slot(str)
    def append_log_message(self, message):
        self.log_text_area.append(message)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
