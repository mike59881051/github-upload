from mailmerge import MailMerge
from datetime import datetime

# ���ɵ��ݺ�ͬ
def GenerateCertify(templateName,newName):
    # ��ģ��
    document = MailMerge(templateName)

    # �滻����
    document.merge(name='����',
                   id='1010101010',
                   year='2020',
                   salary='99999',
                   job='Ƕ��ʽ�����������ʦ')

    # �����ļ�
    document.write(newName)

if __name__ == "__main__":
    templateName = 'н��֤��ģ��.docx'

    # ��ÿ�ʼʱ��
    startTime = datetime.now()

    # ��ʼ����
    for i in range(10000):
        newName = f'./10000��֤��/н��֤��{i}.docx'
        GenerateCertify(templateName,newName)

    # ��ȡ����ʱ��
    endTime = datetime.now()

    # ����ʱ���
    allSeconds = (endTime - startTime).seconds
    print("����10000�ݺ�ͬһ����ʱ: ",str(allSeconds)," ��")

    print("���������")