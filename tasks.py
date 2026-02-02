"""OASP Protocol documentation tasks."""

from invoke.collection import Collection

from scripts.docs.tasks import docs_tasks

ns = Collection()
ns.add_collection(docs_tasks, name="docs")
