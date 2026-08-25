from __future__ import annotations

from urllib.parse import quote

import requests


class SupabaseStoreError(RuntimeError):
    pass


class SupabaseScheduleStore:
    def __init__(self, url: str, key: str, bucket: str = "schedule-files"):
        self.url = url.rstrip("/")
        self.key = key.strip()
        self.bucket = bucket.strip()
        self.headers = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
        }

    def _check(self, response, action):
        if response.ok:
            return response
        detail = response.text[:500]
        raise SupabaseStoreError(f"{action}失败（HTTP {response.status_code}）：{detail}")

    def test_connection(self):
        response = requests.get(
            f"{self.url}/rest/v1/schedules",
            headers=self.headers,
            params={"select": "id", "limit": "1"},
            timeout=15,
        )
        self._check(response, "Supabase 连接测试")
        return True

    def list_schedules(self):
        response = requests.get(
            f"{self.url}/rest/v1/schedules",
            headers=self.headers,
            params={
                "select": "*",
                "status": "neq.archived",
                "order": "created_at.desc",
            },
            timeout=20,
        )
        return self._check(response, "读取排班看板").json()

    def upload_file(self, storage_path: str, content: bytes):
        encoded_path = "/".join(quote(part, safe="") for part in storage_path.split("/"))
        headers = {
            **self.headers,
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "x-upsert": "true",
        }
        response = requests.post(
            f"{self.url}/storage/v1/object/{quote(self.bucket, safe='')}/{encoded_path}",
            headers=headers,
            data=content,
            timeout=60,
        )
        self._check(response, "上传排班文件")

    def download_file(self, storage_path: str):
        encoded_path = "/".join(quote(part, safe="") for part in storage_path.split("/"))
        response = requests.get(
            f"{self.url}/storage/v1/object/authenticated/{quote(self.bucket, safe='')}/{encoded_path}",
            headers=self.headers,
            timeout=60,
        )
        return self._check(response, "下载排班文件").content

    def upsert_schedule(self, metadata: dict):
        headers = {
            **self.headers,
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=representation",
        }
        response = requests.post(
            f"{self.url}/rest/v1/schedules",
            headers=headers,
            params={"on_conflict": "file_hash"},
            json=metadata,
            timeout=20,
        )
        payload = self._check(response, "保存排班记录").json()
        return payload[0] if payload else metadata

    def archive_schedule(self, schedule_id: str):
        headers = {**self.headers, "Content-Type": "application/json", "Prefer": "return=minimal"}
        response = requests.patch(
            f"{self.url}/rest/v1/schedules",
            headers=headers,
            params={"id": f"eq.{schedule_id}"},
            json={"status": "archived"},
            timeout=20,
        )
        self._check(response, "归档排班")

