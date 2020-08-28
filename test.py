import re
from statsimusprime.manager import Manager

m = Manager()
# ss = m.ss_service.service
# l = '1Y0dUyFNv1OZUTSLbNltA1St0m52lRtLfZIfQvhKbUHw'
# s = ss.spreadsheets().get(spreadsheetId = l).execute()
# ll = s['sheets'][0]['properties']['sheetId']
#
# LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
# def to_column(index):
#     if index // 26 == 0:
#         return LETTERS[index]
#     else:
#         return LETTERS[index // 26 - 1] + LETTERS[index % 26]
# def from_column(character_string):
#     if len(character_string) == 1:
#         return LETTERS.index(character_string.upper())
#     elif len(character_string) == 2:
#         return 26*(1 + LETTERS.index(character_string[0].upper())) + LETTERS.index(character_string[1].upper())
#     raise ValueError("Cannot parse column string: `{}`".format(character_string))
#
# parse_cell = re.compile("([A-z]+)([0-9]+)")
# parse_range = re.compile("([A-z]+)([0-9]+)[:]([A-z]+)([0-9]+)")
# def A1_to_bbox(range_string):
#     match = parse_range.match(range_string)
#     if not match is None:
#         r1, c1, r2, c2 = match.groups()
#         r1, r2 = map(from_column, (r1, r2))
#         r2 += 1
#         c1, c2 = map(lambda x: int(x), (c1, c2))
#         c1 -= 1
#     else:
#         r1, c1 = parse_cell.match(range_string).groups()
#         r1 = from_column(r1)
#         r2 = r1 + 1
#         c2 = int(c1)
#         c1 = c2 - 1
#     return (r1, c1, r2, c2)
#
#
#
# def generate_grid_range(sheetId, bbox):
#     return {
#         "sheetId": sheetId,
#         "startRowIndex": bbox[0],
#         "endRowIndex": bbox[2],
#         "startColumnIndex": bbox[1],
#         "endColumnIndex": bbox[3]
#     }
#
# def generate_copy_paste(sheetId, bbox_source, bbox_dest, pasteType = "PASTE_NORMAL"):
#     return {"copyPaste": {
#         "source": generate_grid_range(sheetId, bbox_source),
#         "destination": generate_grid_range(sheetId, bbox_dest),
#         "pasteType": pasteType
#     }}
#
# def generate_add_named_range(sheetId, name, bbox):
#     return {"addNamedRange": {
#         "namedRange": {
#           "name": name,
#           "range": generate_grid_range(sheetId, bbox),
#         }
#       }
#     }
#
# requests = []
# for namedRangeId in [json['namedRangeId'] for json in s['namedRanges'] if json['name'] == 'BOB']:
#     requests.append({"deleteNamedRange" : {
#         "namedRangeId": namedRangeId
#     }})

# request = ss.spreadsheets().batchUpdate(spreadsheetId=l, body={'requests': requests + [
#     generate_copy_paste(
#         sheetId = ll,
#         bbox_source = (0,2,1,3),
#         bbox_dest = (1,2,5,6)
#     )#,
    # generate_add_named_range(
    #     sheetId = ll,
    #     name = "BOB",
    #     bbox = (0,0,2,2)
    # )
# ]}).execute()
# print(request)

# from time import sleep
#
# letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
# citl = [l for l in letters] + ["A"+l for l in letters] + ["B"+l for l in letters] + ["C"+l for l in letters] + ["D"+l for l in letters]
# def get_ranges(spreadsheet):
#     return ["'{0}'!A1:{2}{1}".format(d['properties']['title'], d['properties']['gridProperties']['rowCount'], citl[d['properties']['gridProperties']['columnCount']-1]) for d in spreadsheet['sheets']]
#
# tf = "1sxU3rwzPz-UTNYWJgj50wrJ4ADKsKHuG"
# stats = "1ASFXNteaaUxOZbLGLxGA3BnP4A9GNSB1seFha5_X6oc"
# ssf = "1sc56cHfYU0SoHrgJMw0YYRw1BLTnbdQN"
#
# bu_id = m.drive_service.create_folder(
#     name = "Static Backup",
#     parent_folder_id = tf
# ).get('id')
#
# print("Copying Statistics")
# # Create a stats backup
# bu_stats_id = m.drive_service.copy_to(
#     file_id = stats,
#     name = '_Statistics',
#     destination_folder_id = bu_id
# ).get("id")
#
#
#
# s = m.stats_service.service.spreadsheets().get(spreadsheetId = stats).execute()
# for range in get_ranges(s):
#     data = m.stats_service.get_values(stats, range)
#     print("Copying",range)
#     m.stats_service.update_values(
#         file_id = bu_stats_id,
#         range = data['range'],
#         values = data['values']
#     )
#
# print("Sleeping")
# sleep(100)
#
# for ss in m.drive_service.get_all_children(ssf):
#     if ss['mimeType'] == 'application/vnd.google-apps.spreadsheet':
#         print("Copying",ss['name'])
#
#         # Create a backup
#         bu_ss_id = m.drive_service.copy_to(
#             file_id = ss['id'],
#             name = ss['name'],
#             destination_folder_id = bu_id
#         ).get("id")
#
#         s = m.stats_service.service.spreadsheets().get(spreadsheetId = ss['id']).execute()
#         for range in get_ranges(s):
#             data = m.stats_service.get_values(ss['id'], range)
#             print("Copying",range)
#             m.stats_service.update_values(
#                 file_id = bu_ss_id,
#                 range = data['range'],
#                 values = data['values']
#             )
#
#     print("Sleeping")
#     sleep(100)
