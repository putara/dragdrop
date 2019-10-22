INTDIR = obj
OUTDIR = bin

TARGET_NAME = dragdrop

TARGET_EXE = $(OUTDIR)\$(TARGET_NAME).exe
TARGET_MAP = $(OUTDIR)\$(TARGET_NAME).map
TARGET_PDB = $(OUTDIR)\$(TARGET_NAME).pdb
TARGET_OBJS = \
 "$(INTDIR)\main.obj"\
 "$(INTDIR)\res.res"

CC = cl.exe
LD = link.exe
RC = rc.exe

!if "$(AMD64)" == "1"
TARGET_ARCH = x64
!else
TARGET_ARCH = x86
!endif

CFLAGS = /nologo /c /GF /GL /GR- /GS- /Gy /MD /O1ib2 /W4 /Zi /fp:fast /FAcs /Fa"$(INTDIR)/" /Fo"$(INTDIR)/" /Fd"$(INTDIR)/"
LDFLAGS = /nologo /time /dynamicbase:no /ltcg /machine:$(TARGET_ARCH) /map:"$(TARGET_MAP)" /nxcompat /opt:icf /opt:ref /release /debug /PDBALTPATH:"%_PDB%"
RFLAGS = /d "NDEBUG" /l 0

CDEFS = /D "NDEBUG" /D "_STATIC_CPPLIB" /D "_WINDOWS"
LDLIBS = kernel32.lib user32.lib gdi32.lib comdlg32.lib comctl32.lib shell32.lib shlwapi.lib ole32.lib windowscodecs.lib uxtheme.lib dwmapi.lib

all: "$(INTDIR)" "$(OUTDIR)" "$(TARGET_EXE)"

clean: cleanobj
 -@erase "$(TARGET_EXE)" 2>NUL
 -@erase "$(TARGET_MAP)" 2>NUL
 -@rmdir "$(INTDIR)" 2>NUL

cleanobj: cleanpdb cleanobjonly

cleanpdb:
 -@erase "$(TARGET_PDB)" 2>NUL

cleanobjonly:
 -@erase "$(INTDIR)\*.cod" 2>NUL
 -@erase $(TARGET_OBJS) 2>NUL
 -@erase "$(INTDIR)\vc??.pdb" 2>NUL
 -@erase "$(INTDIR)\vc??.idb" 2>NUL

"$(INTDIR)":
 @if not exist "$(INTDIR)" mkdir "$(INTDIR)"

"$(OUTDIR)":
 @if not exist "$(OUTDIR)" mkdir "$(OUTDIR)"

"$(TARGET_EXE)" : $(TARGET_OBJS)
 $(LD) /out:$@ $(LDFLAGS) $(LDLIBS) $(TARGET_OBJS)

.SUFFIXES: .c .cpp .obj .rc .res

.c{$(INTDIR)}.obj::
 $(CC) $(CFLAGS) $(CDEFS) $<

.cpp{$(INTDIR)}.obj::
 $(CC) $(CFLAGS) $(CDEFS) $<

.rc{$(INTDIR)}.res:
 $(RC) $(RFLAGS) /fo$@ $<

"$(INTDIR)\main.obj": main.cpp
