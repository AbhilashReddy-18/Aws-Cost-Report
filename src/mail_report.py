
import argparse
import json
import os
import re
from html import escape
from pathlib import Path
from urllib.parse import quote

import boto3

from .combine_cost_reports import account_rows, combine, parse_month
from .aws_cost_report import _percentage_change, _shift_month


def email_address(value):
    if not isinstance(value, str) or not re.fullmatch(r'[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+', value):
        raise ValueError('Provide a valid email address without display names or separators.')
    return value


def email_addresses(value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError('Provide at least one recipient email address.')
    recipients = []
    seen = set()
    for item in value.split(','):
        address = email_address(item.strip())
        if address.casefold() not in seen:
            seen.add(address.casefold())
            recipients.append(address)
    return recipients


def settings_from_config(configuration, override=''):
    settings = dict(configuration['notification'])
    for key in ('account_id', 'sns_topic_arn', 'from_email', 'to_email',
                'report_bucket', 'report_prefix', 'kms_key_arn'):
        if not isinstance(settings.get(key), str) or not settings[key].strip():
            raise ValueError(f'Missing notification setting: {key}.')
        settings[key] = settings[key].strip()
    settings['to_email'] = email_addresses(override.strip() or settings['to_email'])
    email_address(settings['from_email'])
    account = settings['account_id']
    if not re.fullmatch(r'[0-9]{12}', account):
        raise ValueError('Notification account_id must have 12 digits.')
    topic = settings['sns_topic_arn'].split(':', 5)
    if (len(topic) != 6 or topic[:3] != ['arn', 'aws', 'sns']
            or not topic[3] or topic[4] != account or not topic[5]):
        raise ValueError('SNS topic must belong to the notification account.')
    settings['region'] = topic[3]
    key = settings['kms_key_arn'].split(':', 5)
    if (len(key) != 6 or key[:3] != ['arn', 'aws', 'kms']
            or key[3] != topic[3] or key[4] != account
            or not key[5].startswith('key/')):
        raise ValueError('KMS key must belong to the notification account and region.')
    prefix = settings['report_prefix'].strip('/')
    if not prefix or '..' in prefix.split('/'):
        raise ValueError('Configure a valid report_prefix.')
    settings['report_prefix'] = prefix
    return settings


def render_email(report, link):
    """Render an overview, monthly totals and account detail using inline styles."""
    def cell(value, heading=False, number=False, background=None):
        tag = 'th' if heading else 'td'
        align = 'right' if number else 'left'
        style = (f'padding:10px;border:1px solid #d9dee5;text-align:{align};'
                 'vertical-align:middle;white-space:nowrap;font-size:13px;')
        if heading:
            style += 'background-color:#193e65;color:#ffffff;font-weight:bold;'
        elif background:
            style += f'background-color:{background};'
        return f'<{tag} style="{style}">{escape(str(value))}</{tag}>'

    def percent_cell(value):
        if value is None:
            return cell('N/A', number=True, background='#f3f6fa')
        color = '#fce8e6' if value > 0 else '#e6f4ea' if value < 0 else '#f3f6fa'
        return cell(f'{value:+.2f}%', number=True, background=color)

    def table(headers, rows):
        return (
            '<table width="100%" cellpadding="0" cellspacing="0" '
            'style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;">'
            '<thead><tr>' + ''.join(cell(v, heading=True) for v in headers)
            + '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table>'
        )

    environments = [name for name in ('Non-Prod', 'Prod')
                    if name in report['environments']]
    monthly_rows = []
    for index, period in enumerate(report['months']):
        total = sum(report['environments'][name]['costs'][index]
                    for name in environments)
        previous = _shift_month(parse_month(period), -12).strftime('%Y-%m')
        prior = sum(data['costs'][previous]
                    for data in report['account_costs'].values())
        cells = cell(period)
        for name in environments:
            cost = report['environments'][name]['costs'][index]
            cells += cell(f'{cost:,.2f}', number=True)
        cells += cell(f'{total:,.2f}', number=True)
        cells += percent_cell(report['total_mom'][index])
        cells += percent_cell(_percentage_change(total, prior))
        monthly_rows.append('<tr>' + cells + '</tr>')
    summary = table(['Month', *environments, 'Total spend', 'MoM %', 'YoY %'],
                    monthly_rows)
    details = []
    account_list = []
    for account in account_rows(report):
        label = (f'{account["account_name"]} ({account["account_id"]})'
                 f' - {account["environment"]}')
        account_list.append(label)
        detail_rows = []
        for period, cost, mom, yoy in zip(report['months'], account['costs'],
                                          account['mom'], account['monthly_yoy']):
            detail_rows.append(
                '<tr>' + cell(period) + cell(f'{cost:,.2f}', number=True)
                + percent_cell(mom) + percent_cell(yoy) + '</tr>'
            )
        details.append(
            '<h3 style="font-size:15px;color:#193e65;margin:22px 0 10px;">'
            + escape(label) + '</h3>'
            + table(['Month', f'Cost ({report["currency"]})', 'MoM %', 'YoY %'],
                    detail_rows)
        )
    report_for = ' & '.join(environments)
    if len(report['account_costs']) > 1:
        report_for += ' (combined)'
    overview_rows = []
    overview = [
        ('Report for', report_for),
        ('Report period', f'{report["months"][0]} to {report["months"][-1]}'),
        ('Cost measure', f'{report["metric"]} - {report["currency"]}'),
        ('Accounts included', '; '.join(account_list)),
    ]
    for label, value in overview:
        overview_rows.append(
            '<tr><td width="150" style="padding:9px 0;font-weight:bold;'
            'vertical-align:top;">' + escape(label) + '</td>'
            '<td style="padding:9px 0;vertical-align:top;">'
            + escape(value) + '</td></tr>'
        )
    overview_rows.append(
        '<tr><td style="padding:9px 0;font-weight:bold;">S3 report location</td>'
        '<td style="padding:9px 0;"><a style="color:#176db0;" href="'
        + escape(link, quote=True) + '">monthly-cost-report.xlsx</a></td></tr>'
    )
    estimates = '; '.join(f'{name}: {", ".join(periods)}'
                          for name, periods in report['estimated_periods'].items() if periods) or 'None flagged'
    html = (
        '<html><body style="margin:0;font-family:Arial,sans-serif;color:#243447;">'
        '<table width="100%" cellpadding="0" cellspacing="0" role="presentation">'
        '<tr><td align="center" style="padding:20px 12px;">'
        '<table width="800" cellpadding="0" cellspacing="0" role="presentation" '
        'style="width:100%;max-width:800px;"><tr><td style="font-family:Arial,sans-serif;">'
        '<h1 style="font-size:26px;text-align:center;color:#193e65;margin:8px 0 24px;">'
        'AWS Monthly Cost Report</h1>'
        '<table width="100%" cellpadding="0" cellspacing="0" role="presentation">'
        '<tr><td style="border-left:5px solid #2282cf;background-color:#f0f6fc;'
        'padding:18px 22px;font-size:14px;">'
        '<h2 style="font-size:20px;color:#193e65;margin:0 0 10px;">Report Overview</h2>'
        '<p style="margin:0 0 12px;">The AWS cost report has been generated successfully. '
        'Below is the spend summary for the selected accounts.</p>'
        '<table width="100%" cellpadding="0" cellspacing="0" role="presentation">'
        + ''.join(overview_rows) + '</table></td></tr></table>'
        '<h2 style="font-size:20px;color:#193e65;margin:26px 0 12px;">'
        f'Monthly Spend Summary ({escape(report["currency"])})</h2>' + summary
        + '<h2 style="font-size:20px;color:#193e65;margin:26px 0 12px;">Account Detail</h2>'
        + ''.join(details)
        + '<p style="font-size:12px;color:#627286;margin:24px 0 10px;">'
        'The Excel link requires AWS access. MoM compares with the previous month; '
        'YoY compares with the same month last year. N/A means the comparison cost was zero. '
        'Red indicates increased cost; green indicates decreased cost. '
        'Costs are not converted between currencies. Only the listed accounts are included.</p>'
        f'<p style="font-size:12px;color:#627286;">Estimated periods: {escape(estimates)}.</p>'
        '</td></tr></table></td></tr></table></body></html>'
    )
    return html.replace('><', '>\n<')


def send_report(session, settings, report, excel_path, run_id):
    key = (f'{settings["report_prefix"]}/{report["report_month"]}/'
           f'build-{run_id}/monthly-cost-report.xlsx')
    link = (f'https://s3.console.aws.amazon.com/s3/object/{settings["report_bucket"]}'
            f'?region={quote(settings["region"])}&prefix={quote(key, safe="")}')
    body = render_email(report, link)
    # Publish one request containing all validated recipients.
    content = {
        'event_type': 'push',
        'push': {
            'notification_channel': 'email',
            'email': {
                'from': settings['from_email'],
                'to': ','.join(settings['to_email']),
                'subject': f"AWS Cost Report - {report['report_month']}",
                'body': body,
            },
        },
    }
    message = json.dumps({'default': json.dumps(content)})
    if len(message.encode('utf-8')) > 262144:
        raise ValueError('Email payload exceeds the SNS message size limit.')
    messages = [message]
    actual = session.client('sts', region_name=settings['region']).get_caller_identity()['Account']
    if actual != settings['account_id']:
        raise ValueError(f'Notification credentials belong to {actual}; expected {settings["account_id"]}.')
    if not excel_path.is_file() or excel_path.stat().st_size > 20 * 1024 * 1024:
        raise ValueError('Excel report must exist and be no larger than 20 MiB.')
    with excel_path.open('rb') as source:
        session.client('s3', region_name=settings['region']).put_object(
            Bucket=settings['report_bucket'], Key=key, Body=source,
            ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ServerSideEncryption='aws:kms', SSEKMSKeyId=settings['kms_key_arn'],
        )
    sns = session.client('sns', region_name=settings['region'])
    notifications = []
    for message in messages:
        response = sns.publish(
            TopicArn=settings['sns_topic_arn'], Message=message, MessageStructure='json',
        )
        notifications.append(response['MessageId'])
        print(f'Email request {len(notifications)}/{len(messages)} published; '
              f'message ID: {response["MessageId"]}.')
    return {'status': 'published', 'message_ids': notifications,
            'report_s3_key': key}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=Path('reporting-config.json'))
    parser.add_argument('--input', type=Path, default=Path('report-data'))
    parser.add_argument('--excel', type=Path,
                        default=Path('report-output-combined/monthly-cost-report.xlsx'))
    parser.add_argument('--accounts', nargs='+', required=True)
    parser.add_argument('--run-id', required=True)
    args = parser.parse_args()
    if not re.fullmatch(r'[0-9]+', args.run_id):
        parser.error('--run-id must be a Jenkins build number.')
    configuration = json.loads(args.config.read_text(encoding='utf-8'))
    settings = settings_from_config(configuration, os.environ.get('REPORT_EMAIL_TO', ''))
    records = [json.loads(path.read_text(encoding='utf-8'))
               for path in sorted(args.input.glob('*.json'))]
    report = combine(records, args.accounts)
    result = send_report(boto3.Session(), settings, report, args.excel, args.run_id)
    print(f'Published {len(result["message_ids"])} email requests. Check each recipient inbox.')


if __name__ == '__main__':
    main()
