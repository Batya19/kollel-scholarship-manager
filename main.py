from ui.main_window import ModernKollelUI
from splash_screen import SplashScreen

def main():
    splash = SplashScreen()
    splash.run_with_app(lambda: ModernKollelUI().run())

if __name__ == "__main__":
    main()