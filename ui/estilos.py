# ui/estilos.py
"""
Estilos CSS de PyQt5 para la aplicación.

Proporciona temas claro y oscuro con el color corporativo #16A085 (verde turquesa).
"""


class Estilos:
    """
    Clase con estilos CSS para temas claro y oscuro de PyQt5.
    
    Color principal: #16A085 (verde turquesa)
    """
    
    @staticmethod
    def obtener_estilo(tema):
        """
        Obtiene el estilo CSS según el tema seleccionado.
        
        Args:
            tema (str): 'light' para tema claro, 'dark' para tema oscuro
            
        Returns:
            str: String con el CSS de PyQt5
        """
        if tema == "dark":
            return Estilos.estilo_oscuro()
        else:
            return Estilos.estilo_claro()
    
    @staticmethod
    def estilo_claro():
        """
        Retorna el estilo CSS para el tema claro.
        
        Características:
        - Fondo: #F0F8F5 (verde muy claro)
        - Controles: Blancos con bordes grises
        - Color primario: #16A085
        
        Returns:
            str: CSS de PyQt5 para tema claro
        """
        return """
            QMainWindow {
                background-color: #F0F8F5;
            }
            QWidget {
                background-color: #F0F8F5;
                color: #2C3E50;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #16A085;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #16A085;
            }
            QPushButton {
                background-color: #16A085;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #138D75;
            }
            QPushButton:pressed {
                background-color: #0E6655;
            }
            QPushButton:disabled {
                background-color: #BDC3C7;
            }
            QLineEdit, QTextEdit, QDateEdit {
                padding: 8px;
                border: 2px solid #BDC3C7;
                border-radius: 5px;
                background-color: white;
                selection-background-color: #16A085;
            }
            QLineEdit:focus, QTextEdit:focus {
                border: 2px solid #16A085;
            }
            QTabWidget::pane {
                border: 2px solid #16A085;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #ECF0F1;
                color: #2C3E50;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background-color: #16A085;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #BDC3C7;
            }
            QProgressBar {
                border: 2px solid #BDC3C7;
                border-radius: 5px;
                text-align: center;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #16A085;
                border-radius: 5px;
            }
            QTreeWidget {
                border: 2px solid #BDC3C7;
                border-radius: 5px;
                background-color: white;
            }
            QTreeWidget::item:selected {
                background-color: #16A085;
                color: white;
            }
        """
    
    @staticmethod
    def estilo_oscuro():
        """
        Retorna el estilo CSS para el tema oscuro.
        
        Características:
        - Fondo: #1E1E1E (negro suave)
        - Controles: Grises oscuros
        - Color primario: #16A085
        
        Returns:
            str: CSS de PyQt5 para tema oscuro
        """
        return """
            QMainWindow {
                background-color: #1E1E1E;
            }
            QWidget {
                background-color: #1E1E1E;
                color: #E0E0E0;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #16A085;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #2D2D2D;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #16A085;
            }
            QPushButton {
                background-color: #16A085;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #138D75;
            }
            QPushButton:pressed {
                background-color: #0E6655;
            }
            QPushButton:disabled {
                background-color: #404040;
            }
            QLineEdit, QTextEdit, QDateEdit {
                padding: 8px;
                border: 2px solid #404040;
                border-radius: 5px;
                background-color: #2D2D2D;
                color: #E0E0E0;
                selection-background-color: #16A085;
            }
            QLineEdit:focus, QTextEdit:focus {
                border: 2px solid #16A085;
            }
            QTabWidget::pane {
                border: 2px solid #16A085;
                border-radius: 8px;
                background-color: #2D2D2D;
            }
            QTabBar::tab {
                background-color: #404040;
                color: #E0E0E0;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background-color: #16A085;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #505050;
            }
            QProgressBar {
                border: 2px solid #404040;
                border-radius: 5px;
                text-align: center;
                background-color: #2D2D2D;
                color: #E0E0E0;
            }
            QProgressBar::chunk {
                background-color: #16A085;
                border-radius: 5px;
            }
            QTreeWidget {
                border: 2px solid #404040;
                border-radius: 5px;
                background-color: #2D2D2D;
                color: #E0E0E0;
            }
            QTreeWidget::item:selected {
                background-color: #16A085;
                color: white;
            }
        """