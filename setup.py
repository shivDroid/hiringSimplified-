import subprocess
import sys
import os

def check_python_version():
    """Check if Python version is 3.7 or higher"""
    if sys.version_info < (3, 7):
        print("Error: Python 3.7 or higher is required")
        sys.exit(1)

def create_virtual_environment():
    """Create a virtual environment"""
    if not os.path.exists('venv'):
        subprocess.run([sys.executable, '-m', 'venv', 'venv'])
        print("Virtual environment created successfully")
    else:
        print("Virtual environment already exists")

def install_requirements():
    """Install required packages"""
    if os.name == 'nt':  # Windows
        pip_path = os.path.join('venv', 'Scripts', 'pip')
    else:  # Linux/Mac
        pip_path = os.path.join('venv', 'bin', 'pip')
    
    subprocess.run([pip_path, 'install', '-r', 'requirements.txt'])
    print("Requirements installed successfully")

def main():
    print("Setting up the email automation system...")
    check_python_version()
    create_virtual_environment()
    install_requirements()
    print("\nSetup completed successfully!")
    print("\nNext steps:")
    print("1. Place your Google credentials.json file in the project directory")
    print("2. Activate the virtual environment:")
    if os.name == 'nt':  # Windows
        print("   venv\\Scripts\\activate")
    else:  # Linux/Mac
        print("   source venv/bin/activate")
    print("3. Run the main script: python main.py")

if __name__ == "__main__":
    main()