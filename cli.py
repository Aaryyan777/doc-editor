import typer
from rich.console import Console
from rich.table import Table
import os
import pdf_ops

app = typer.Typer(help="Doc-Tor: A powerful CLI PDF Editor built with Python.")
console = Console()

@app.command()
def info(file_path: str):
    """Show metadata and information about a PDF file."""
    try:
        data = pdf_ops.get_pdf_info(file_path)
        table = Table(title=f"PDF Metadata: {os.path.basename(file_path)}")
        table.add_column("Property", style="cyan", no_wrap=True)
        table.add_column("Value", style="magenta")
        for k, v in data.items():
            table.add_row(k.replace("_", " ").title(), str(v))
        console.print(table)
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def extract(file_path: str, output: str = typer.Option(None, help="Output text file path")):
    """Extract text from a PDF file."""
    try:
        with console.status("[bold green]Extracting text..."):
            saved_path = pdf_ops.extract_text_from_pdf(file_path, output)
        console.print(f"[bold green]Success![/bold green] Text extracted to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def merge(files: list[str], output: str = typer.Option("merged.pdf", help="Output PDF file path")):
    """Merge multiple PDF files into one."""
    if len(files) < 2:
        console.print("[bold red]Error:[/bold red] Please provide at least two files.")
        return
    try:
        with console.status("[bold green]Merging files..."):
            saved_path = pdf_ops.merge_pdfs(files, output)
        console.print(f"[bold green]Success![/bold green] Merged to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def redact(file_path: str, text: str, output: str = typer.Option(None, help="Output PDF file path")):
    """Redact (black out) specific text in the PDF."""
    try:
        with console.status(f"[bold green]Redacting '{text}'..."):
            count, saved_path = pdf_ops.redact_pdf(file_path, text, output)
        if count == 0:
            console.print(f"[yellow]No instances of '{text}' found.[/yellow]")
        else:
            console.print(f"[bold green]Success![/bold green] Redacted {count} occurrences. Saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def edit(file_path: str, old_text: str, new_text: str, output: str = typer.Option(None, help="Output PDF file path")):
    """Experimental: Search and replace text."""
    try:
        with console.status(f"[bold green]Replacing '{old_text}'..."):
            count, saved_path = pdf_ops.edit_pdf_text(file_path, old_text, new_text, output)
        console.print(f"[bold green]Success![/bold green] Replaced {count} occurrences. Saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def pdf_to_word(file_path: str, output: str = typer.Option(None, help="Output Word (.docx) file path")):
    """Convert PDF to a Word Document (.docx) for full editing."""
    try:
        with console.status("[bold green]Converting PDF to Word..."):
            saved_path = pdf_ops.convert_pdf_to_word(file_path, output)
        console.print(f"[bold green]Success![/bold green] Converted to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def word_to_pdf(file_path: str, output: str = typer.Option(None, help="Output PDF file path")):
    """Convert a Word Document (.docx) back to PDF."""
    try:
        with console.status("[bold green]Converting Word to PDF..."):
            saved_path = pdf_ops.convert_word_to_pdf(file_path, output)
        console.print(f"[bold green]Success![/bold green] Converted to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def rotate(file_path: str, degrees: int = typer.Option(90, help="Rotation angle (90, 180, 270)"), output: str = typer.Option(None)):
    """Rotate all pages in the PDF."""
    try:
        saved_path = pdf_ops.rotate_pages(file_path, degrees, output)
        console.print(f"[bold green]Success![/bold green] Rotated by {degrees} degrees. Saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def delete_pages(file_path: str, pages: str = typer.Option(..., help="Pages to delete (e.g. '1,3,5')"), output: str = typer.Option(None)):
    """Delete specific pages from the PDF."""
    try:
        page_list = [int(p.strip()) - 1 for p in pages.split(",")] # Convert to 0-based
        saved_path = pdf_ops.delete_pages(file_path, page_list, output)
        console.print(f"[bold green]Success![/bold green] Deleted pages. Saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def split(file_path: str, start: int, end: int, output: str = typer.Option(None)):
    """Extract a range of pages to a new file."""
    try:
        saved_path = pdf_ops.extract_page_range(file_path, start, end, output)
        console.print(f"[bold green]Success![/bold green] Extracted pages {start}-{end}. Saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def encrypt(file_path: str, password: str, output: str = typer.Option(None)):
    """Protect PDF with a password."""
    try:
        saved_path = pdf_ops.encrypt_pdf(file_path, password, output)
        console.print(f"[bold green]Success![/bold green] Encrypted file saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

@app.command()
def decrypt(file_path: str, password: str, output: str = typer.Option(None)):
    """Remove password protection from PDF."""
    try:
        saved_path = pdf_ops.decrypt_pdf(file_path, password, output)
        console.print(f"[bold green]Success![/bold green] Decrypted file saved to '{saved_path}'")
    except Exception as e:
        console.print(f"[bold red]Error:[/bold red] {e}")

if __name__ == "__main__":
    app()