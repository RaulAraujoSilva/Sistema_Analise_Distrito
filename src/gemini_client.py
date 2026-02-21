# -*- coding: utf-8 -*-
"""
Wrapper para a API do Gemini com rate limiting e retry.
Suporta: análise de texto+imagens, upload de PDF, e geração de imagens.
"""
import os
import shutil
import tempfile
import time
import logging
from pathlib import Path

from google import genai
from google.genai import types

logger = logging.getLogger(__name__)


class GeminiAuditClient:
    MODEL = "gemini-3-flash-preview"
    IMAGE_MODEL = "gemini-3-pro-image-preview"
    MAX_RETRIES = 3
    BASE_DELAY = 5        # segundos entre chamadas (Flash tem limites mais generosos)
    RETRY_BACKOFF = 15    # segundos após rate limit

    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)
        self._pdf_file = None

    def upload_pdf(self, pdf_path: str) -> None:
        """Upload da apostila PDF via File API (reutilizável por 48h)."""
        logger.info(f"Enviando PDF para File API: {pdf_path}")
        tmp_dir = tempfile.mkdtemp()
        safe_name = "apostila_abar.pdf"
        safe_path = os.path.join(tmp_dir, safe_name)
        shutil.copy2(pdf_path, safe_path)
        logger.info(f"  Cópia temporária: {safe_path}")
        try:
            self._pdf_file = self.client.files.upload(file=safe_path)
            logger.info(f"PDF enviado. Name: {self._pdf_file.name}, URI: {self._pdf_file.uri}")
        finally:
            try:
                os.remove(safe_path)
                os.rmdir(tmp_dir)
            except OSError:
                pass

    def analyze_section(
        self,
        system_prompt: str,
        section_prompt: str,
        image_paths: list,
        thinking_level: str = "medium",
        include_pdf: bool = False,
    ) -> str:
        """
        Envia análise com texto + imagens (e opcionalmente PDF) ao Gemini.
        Retorna o texto gerado.
        """
        contents = []

        # 1. PDF (apenas se solicitado e disponível)
        if include_pdf and self._pdf_file:
            contents.append(self._pdf_file)

        # 2. Imagens PNG inline
        for img_path in image_paths:
            p = Path(img_path)
            if p.exists():
                contents.append(
                    types.Part.from_bytes(
                        data=p.read_bytes(),
                        mime_type="image/png",
                    )
                )
                logger.debug(f"  Imagem anexada: {p.name} ({p.stat().st_size / 1024:.0f} KB)")
            else:
                logger.warning(f"  Imagem não encontrada: {img_path}")

        # 3. Texto do prompt (system + section combinados)
        contents.append(f"{system_prompt}\n\n---\n\n{section_prompt}")

        # 4. Chamada com retry
        for attempt in range(self.MAX_RETRIES):
            try:
                logger.info(f"  Chamando Gemini (tentativa {attempt + 1}/{self.MAX_RETRIES})...")
                response = self.client.models.generate_content(
                    model=self.MODEL,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        thinking_config=types.ThinkingConfig(
                            thinking_level=thinking_level,
                        ),
                    ),
                )

                text = response.text or ""
                logger.info(f"  Resposta recebida: {len(text)} caracteres")
                time.sleep(self.BASE_DELAY)
                return text

            except Exception as e:
                err_str = str(e)
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    wait = self.RETRY_BACKOFF * (attempt + 1)
                    logger.warning(f"  Rate limit atingido. Aguardando {wait}s...")
                    time.sleep(wait)
                else:
                    logger.error(f"  Erro na API: {e}")
                    if attempt == self.MAX_RETRIES - 1:
                        raise
                    time.sleep(self.RETRY_BACKOFF)

        raise RuntimeError(f"Máximo de tentativas ({self.MAX_RETRIES}) excedido")

    def generate_image(self, prompt: str, output_path: str) -> bool:
        """
        Gera uma imagem usando Gemini Image Generation.
        Retorna True se a imagem foi salva com sucesso.
        """
        for attempt in range(self.MAX_RETRIES):
            try:
                logger.info(f"  Gerando imagem (tentativa {attempt + 1}/{self.MAX_RETRIES})...")
                response = self.client.models.generate_content(
                    model=self.IMAGE_MODEL,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_modalities=["TEXT", "IMAGE"],
                    ),
                )

                for part in response.candidates[0].content.parts:
                    if part.inline_data and part.inline_data.mime_type.startswith("image/"):
                        image_data = part.inline_data.data
                        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
                        with open(output_path, "wb") as f:
                            f.write(image_data)
                        logger.info(f"  Imagem salva: {output_path} ({len(image_data) / 1024:.0f} KB)")
                        time.sleep(self.BASE_DELAY)
                        return True

                logger.warning("  Nenhuma imagem na resposta do Gemini.")
                time.sleep(self.BASE_DELAY)
                return False

            except Exception as e:
                err_str = str(e)
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    wait = self.RETRY_BACKOFF * (attempt + 1)
                    logger.warning(f"  Rate limit. Aguardando {wait}s...")
                    time.sleep(wait)
                else:
                    logger.error(f"  Erro na geração de imagem: {e}")
                    if attempt == self.MAX_RETRIES - 1:
                        return False
                    time.sleep(self.RETRY_BACKOFF)

        return False
