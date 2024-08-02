# https://boto3.amazonaws.com/v1/documentation/api/latest/index.html

import boto3
import pandas as pd
import os
from openpyxl import load_workbook

session = boto3.Session(profile_name='config-stage')
file_name = 'target_groups_stage-ext.xlsx'
as_is_version = '-v1.23'
to_be_version = '-v1.24'

def get_excel(sheet_name):
  sheet_name = sheet_name + as_is_version

  if not os.path.exists(file_name):
    print(f"### File does not exist - {file_name} ")
    return None

  try:
    df = pd.read_excel(sheet_name=sheet_name)
    return df.to_dict(orient='records')
  except Exception as e:
    print(f"!! Error file is {file_name}, sheet is {sheet_name} msg : {e}")
    return None

def excel_input(data_frame, sheet_name):
  data_frame = pd.DataFrame(data_frame)
  sheet_name = sheet_name + as_is_version

  if not sheet_name:
    sheet_name = 'none'

  if os.path.exists(file_name):
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
      try:
        existing_data = pd.read_excel(file_name, sheet_name=sheet_name)
        startrow = len(existing_data) + 1
      except ValueError:
        startrow = 0
      data_frame.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=not bool(startrow))
  else:
    data_frame.to_excel(file_name, sheet_name=sheet_name, index=False)

  print(f'### save to {file_name}, sheet-name is {sheet_name}')

def get_instances(instance_name):
  client = session.client('ec2')

  response = client.describe_instances(
    Filters=[
      {
        'Name': 'tag:Name',
        'Values': [instance_name]
      },
    ]
  )

  instance_info = []
  for reservation in response['Reservations']:
    for instance in reservation['Instances']:
      if instance['State']['Name'] == 'terminated':
        continue
      instance_id = instance['InstanceId']
      instance_name = None
      availability_zone = instance['Placement']['AvailabilityZone']
      instance_type = instance['InstanceType']
      print(f'availability_zone : {availability_zone}, instance_type : {instance_type}')
      for tag in instance.get('Tags', []):
        if tag['Key'] == 'Name':
          instance_name = tag['Value']
          break
      instance_info.append({
        'instance_id': instance_id,
        'instance_name': instance_name,
        'availability_zone': availability_zone,
        'instance_type': instance_type
      })

  return instance_info

def get_new_instances(instance_name, as_is_instances):
  get_all_instances = get_instances(instance_name)
  to_be_instances = []

  for all_instance in get_all_instances:
    found = False
    for instance in as_is_instances:
      if all_instance['instance_id'] == instance['instance_id']:
        found = True
        break
    if not found:
      to_be_instances.append(all_instance)
  return to_be_instances

def register_targets(as_is_target_group, to_be_instances):
  client = session.client('elbv2')
  for tg in as_is_target_group:
    tg_arn = tg['TargetGroupArn']
    tg_port = tg['Port']
    for instance in to_be_instances:
      instance_id = instance['instance_id']
      target = {
        'Id': instance_id,
        'Port': tg_port
      }
      if 'availability_zone' in instance and tg.get('instance_type') == 'ip':
        target['AvailabilityZone'] = instance['availability_zone']

      response = client.register_targets(
          TargetGroupArn=tg_arn,
          Targets=[target]
      )

      print(f'register target ok - instance_id : {instance_id}, port : {tg_port}')

def get_target_groups(to_be_instances):
  instance_ids = [instance['instance_id'] for instance in to_be_instances]

  client = session.client('elbv2')
  response = client.describe_target_groups()

  target_groups = []
  for tg in response['TargetGroups']:
    tg_arn = tg['TargetGroupArn']
    tg_name = tg['TargetGroupName']
    tg_port = tg['Port']
    tg_type = tg['TargetType']
    tg_lb_arn = tg['LoadBalancerArns']
    try:
      targets_response = client.describe_target_health(TargetGroupArn=tg_arn)
      for target in targets_response['TargetHealthDescriptions']:
        instance_id = target['Target']['Id']
        if instance_id in instance_ids:
          target_groups.append({
            'TargetGroupArn' : tg_arn
            , 'AlbArn' : tg_lb_arn
            , 'Name' : tg_name
            , 'Port' : tg_port
            , 'InstanceId': instance_id
            , 'Type' : tg_type
            , 'Note' : 'Add'
          })

    except Exception as e:
      print(f"!! An error target - {tg_arn} : {e}")
      continue
  return target_groups

def main():
  print('====== Register target_group start ======')

  as_is_instances = get_excel('AsIs-instance')
  as_is_target_group = get_excel('AsIs-TargetGroups')

  instance_name = 'stg-ext-ng'
  to_be_instances = get_new_instances(instance_name, as_is_instances)
  print(f'### Get TO Be Instances Info : {to_be_instances}')
  excel_input(to_be_instances, 'ToBe-instance')

  register_response = register_targets(as_is_target_group, to_be_instances)

  to_be_target_group = get_target_groups(to_be_instances)
  excel_input(to_be_target_group, 'ToBe-TargetGroups')

if __name__ == "__main__":
    main()

