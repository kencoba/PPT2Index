@rem PowerPoint�t�@�C������A���������쐬����
@rem �g�p�@
@rem PPT2Index.bat PowerPoint�t�@�C�� �e�L�X�g�f�[�^���o���ʃt�@�C�� �����p�L�[���[�h�t�@�C��
@rem ex: > PPT2Index.bat CourseWare.pptx output.xml keywords.txt

cscript PPT2Text.vbs %1 //nologo > %2 
scala makeIndex.scala %2 %3