# https://boto3.amazonaws.com/v1/documentation/api/latest/index.html

import boto3
import pandas as pd
import os
from openpyxl import load_workbook

session = boto3.Session(profile_name='config-stage')
file_name = 'target_groups_stage-ext.xlsx'
version = '-v1.23'

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
  print(f'############ yujin : {response}')
  instance_info = []
  for reservation in response['Reservations']:
    for instance in reservation['Instances']:
      if instance['State']['Name'] == 'terminated':
        break
      instance_id = instance['InstanceId']
      instance_name = None
      for tag in instance.get('Tags', []):
        if tag['Key'] == 'Name':
          instance_name = tag['Value']
          break
      instance_info.append({
        'instance_id': instance_id,
        'instance_name': instance_name
      })

  return instance_info

def get_target_groups(eks_instance):
  instance_ids = [instance['instance_id'] for instance in eks_instance]

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
          })

    except Exception as e:
      print(f"!! An error target - {tg_arn} : {e}")
      continue
  return target_groups

def excel_input(data_frame, sheet_name):
  data_frame = pd.DataFrame(data_frame)
  sheet_name = sheet_name + version
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


def main():
  print('====== target_group start ======')
  instance_name = 'stg-ext-ng'
  as_is_instances = get_instances(instance_name)
  print(f'### Get Instances Info : {as_is_instances}')

  if not as_is_instances:
    print("!! No Instance !!")
    return
  excel_input(as_is_instances, 'AsIs-instance')

  target_groups = get_target_groups(as_is_instances)
  print(f'### Target Group Info : {target_groups}')
  if not target_groups:
      print("!! No target group !!")
      return

  excel_input(target_groups, 'AsIs-TargetGroups')
  print("### File save fin ###")

if __name__ == "__main__":
    main()
