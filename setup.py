from setuptools import setup, find_packages

setup(
    name="delivery-plan-processor",
    version="0.1.0",
    description="Excel到货计划处理工具",
    author="Rien",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.3.0",
        "openpyxl>=3.0.7",
        "PyYAML>=5.4.1",
        "pywin32>=300",
        "python-dotenv>=0.19.0",
    ],
    python_requires=">=3.8",
    entry_points={
        'console_scripts': [
            'delivery-plan=main:local_handler',
        ],
    },
) 