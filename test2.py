import re
from statsimusprime.manager import Manager, MediaFileUpload

m = Manager()
ss = m.ss_service.service
ds = m.drive_service.service
mime = "application/json"#"application/vnd.google-apps.script+json"

id = "1zk_lboCbmlGhs6I6m_wZf0N-QB5AB9-9"

file_md = {
    "name" : "env",
    "mimeType" : mime,
    'parents' : [m.env['top_folder_id']]
}

media = MediaFileUpload(
    "..env",
    mimetype = "application/json",
    resumable = True
)

# ds.files().create(
#     body=file_md,
#     media_body=media,
#     fields='id'
# ).execute()

ds.files().update(
    fileId = id,
    # body=file,
    media_body=media,
    fields='id'
).execute()
