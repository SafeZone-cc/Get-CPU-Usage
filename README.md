# Get CPU Usage

�㤥� ������� ��� ����ண� ��⥪� bitcoin-�����஢ ��� ��㣨� "���������" ����ᮢ.

�����뢠�� ⠪�� ���ଠ��:

* ��� �����
* Process ID
* ����
* ��� ᮮ⢥�����饩 �㦡� (�ᥢ����� + ���ᠭ��)
* �६� CPU
* % ����㧪�:

## ����
- ⥪��� ����㧪� �� ��� � 2 ᥪ. - � ���� **ProcessCPU_Current.csv**
- �।��� ����㧪� �� �� �६� ࠡ��� �� - � ���� **ProcessCPU_Average.csv**

## ����㧪� �������� �� ��㫥
(����� �६��� KernelModeTime + UserModeTime �����)
`/`
(����� �६��� KernelModeTime + UserModeTime ��⥬� � 楫��)
`* 100`

���ଠ�� ������ �� ��ꥪ� WMI (Win32_Process, Win32_Service)


## �������� �� �ᯮ�짮�����:
1. ��ᯠ��� ��娢.
2. ������� 䠩� GetCPUUsage.vbs
�᫨ ����� ᮮ�饭�� �� User Accaunt Control, �⢥砥� "��".
4. ��������, ���� �� ����� ᮮ�饭�� "��⮢�."
5. �뫮��� � ⥬�, ��� ��� ����뢠�� ������, 䠩��:
- ProcessCPU_Current.csv
- ProcessCPU_Average.csv

㯠����� �� � ��娢 �ଠ� zip.

�᫨ �ந��諠 �訡��, �뫮��� �� [�ਭ��](http://safezone.cc/threads/kak-sdelat-skrinshot.21063/) � [�⮩ ⥬�](http://safezone.cc/threads/getcpuusage.23005/).

�᫨ ���� �� ������, ������ ����� �����襩 ��� �� ��������� ����� � ��⥬ ������� F5.
