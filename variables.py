import yaml as ym
from yaml.loader import SafeLoader
import os, fnmatch


class Var:
    with open('dependencies\config.yaml') as f:
        configYaml = ym.load(f, Loader=SafeLoader)
    rodape =f"""
            <p><span style="color:#3366cc">Atenciosamente</span><span style="color:#3399cc">,</span><br />
                <strong><span style="color:#3366cc">{configYaml['config']['email_signature']['name']}</span></strong><br />
                <span style="color:#3366cc">{configYaml['config']['email_signature']['job']}<br />
                {configYaml['config']['email_signature']['phone_number']}</span><br />
                <a href="mailto:{configYaml['config']['email_signature']['email']}">{configYaml['config']['email_signature']['email']}</a></p>
            """

    