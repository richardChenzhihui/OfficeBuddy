"""Allow running office_agent as a module: python -m office_agent"""
import sys
from pathlib import Path

# Add parent directory to Python path
current_dir = Path(__file__).parent.absolute()
parent_dir = current_dir.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

# Import and run test
if __name__ == "__main__":
    from office_agent.test_agent import *
    print("\nRun specific tests:")
    print("  python -m office_agent.test_validation")
    print("  python -m office_agent.test_functional")
    print("  python -m office_agent.test_agent")
