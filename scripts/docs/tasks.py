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
def deploy(c: Context, versions: str = "latest") -> None:
    """Deploy documentation to the documentation server.

    Args:
        versions: Comma-separated version labels to deploy (default: latest).
                  Use for bug fixes that need to update historical versions.

    Examples:
        inv docs.deploy                      # Deploy only 'latest'
        inv docs.deploy --versions=latest    # Same as above
        inv docs.deploy --versions=0.1.0,latest  # Bug fix: update both versions
    """
    version_list = [v.strip() for v in versions.split(",")]

    for version in version_list:
        print(f"Deploying version: {version}")
        c.run(f"mike deploy --push {version}")


@task
def serve_versioned(c: Context) -> None:
    """Serve versioned documentation locally."""
    c.run("mike serve")


docs_tasks = Collection("docs")
docs_tasks.add_task(serve)
docs_tasks.add_task(build)
docs_tasks.add_task(deploy)
docs_tasks.add_task(serve_versioned, name="serve-versioned")
