"""Documentation build and deployment tasks."""

from invoke.collection import Collection
from invoke.context import Context
from invoke.tasks import task


@task
def serve(c: Context) -> None:
    """Start MkDocs development server with live reload."""
    c.run("mkdocs serve")


@task
def build(c: Context) -> None:
    """Build static documentation."""
    c.run("mkdocs build")


@task
def deploy(c: Context, version: str = "latest") -> None:
    """Deploy documentation to the documentation server.

    Args:
        version: Version label for the deployment (default: latest)
    """
    # Build with mike for versioning
    c.run(f"mike deploy --push --update-aliases {version}")


@task
def serve_versioned(c: Context) -> None:
    """Serve versioned documentation locally."""
    c.run("mike serve")


docs_tasks = Collection("docs")
docs_tasks.add_task(serve)
docs_tasks.add_task(build)
docs_tasks.add_task(deploy)
docs_tasks.add_task(serve_versioned, name="serve-versioned")
