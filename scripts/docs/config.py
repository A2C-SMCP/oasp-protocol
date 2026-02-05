"""文档部署配置管理。

从环境变量加载部署配置，与 TFRobotV2 保持一致。
"""

import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class DocServerConfig:
    """文档服务器配置。

    Attributes:
        host: 服务器主机地址或 IP
        port: SSH 端口，默认 22
        user: SSH 用户名，默认 root
        password: SSH 密码（可选，优先使用密钥）
        key_filename: SSH 私钥文件路径
        deploy_path: 文档部署路径
    """

    host: str
    port: int = 22
    user: str = "root"
    password: Optional[str] = None
    key_filename: Optional[str] = None
    deploy_path: str = "/var/www/doc.turingfocus.cn/oasp"


@dataclass
class DeployConfig:
    """部署配置总入口。

    Attributes:
        server: 文档服务器配置
    """

    server: DocServerConfig

    @classmethod
    def from_env(cls) -> "DeployConfig":
        """从环境变量加载配置。

        环境变量与 TFRobotV2 保持一致：
        - DOCS_SERVER_HOST: 服务器地址
        - DOCS_SERVER_PORT: SSH 端口
        - DOCS_SERVER_USER: SSH 用户名
        - DOCS_SERVER_PASSWORD: SSH 密码
        - DOCS_SERVER_KEY_FILE: SSH 私钥文件路径

        Returns:
            DeployConfig: 加载的配置对象
        """
        server = DocServerConfig(
            host=os.getenv("DOCS_SERVER_HOST") or "",
            port=int(os.getenv("DOCS_SERVER_PORT") or "22"),
            user=os.getenv("DOCS_SERVER_USER") or "root",
            password=os.getenv("DOCS_SERVER_PASSWORD") or None,
            key_filename=os.getenv("DOCS_SERVER_KEY_FILE") or None,
            deploy_path=os.getenv("DOCS_DEPLOY_PATH") or "/var/www/doc.turingfocus.cn/oasp",
        )

        return cls(server=server)

    def validate(self) -> list[str]:
        """验证配置是否完整。

        Returns:
            list[str]: 错误信息列表，空列表表示验证通过
        """
        errors = []

        if not self.server.host:
            errors.append("DOCS_SERVER_HOST 未设置")

        if not self.server.password and not self.server.key_filename:
            errors.append("DOCS_SERVER_PASSWORD 或 DOCS_SERVER_KEY_FILE 至少需要设置一个")

        return errors
