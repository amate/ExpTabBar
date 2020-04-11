// @file	OleDragDropTabCtrl.h
// @brief	MTL : �^�u�ł̃h���b�O���h���b�v����

#pragma once

#include <vector>
#include "TabCtrlImpl.h"
#include "MtlDragDrop.h"
#include "DataObject.h"


using namespace std;

namespace MTL {


// for debug
#ifdef _DEBUG
	const bool _Mtl_DragDrop_TabCtrl2_traceOn = false;
	#define DDTCTRACE	if (_Mtl_DragDrop_TabCtrl2_traceOn) ATLTRACE
#else
	#define DDTCTRACE
#endif


#define GetPIDLFolder(pida) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[0])
	#define GetPIDLItem(pida, i) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[i+1])

///////////////////////////////////////////////////////////////////////
// OleDragDropTabCtrl

template <class T>
class COleDragDropTabCtrl
	: public CTabCtrlImpl
	, public IDropTargetImpl<T>
	, public IDropSourceImpl<T>
{
public:
	DECLARE_WND_CLASS_EX(_T("Mtl_OleDragDrop_TabCtrl"), CS_DBLCLKS, COLOR_BTNFACE)	// don't forget CS_DBLCLKS

private:
	// Data members
	CRect				m_rcInvalidateOnDrawingInsertionEdge;
	CSimpleArray<int>	m_arrCurDragItems;	// ����Drag����Ă���^�u�̃C���f�b�N�X

//	CImageList			m_DragImage;

protected:
	bool				m_bDragFromItself;	// �^�u�o�[�̃^�u��Drag&Drop����Ă���Ȃ�true
	bool				m_bLeftButton;		// ���{�^������Drag���Ă�Ȃ�true
	CString				m_strDraggingDrive;	// ���݃h���b�O���̃t�@�C���̃h���C�u(C:\�Ƃ�)

public:
	// �R���X�g���N�^
	COleDragDropTabCtrl() : m_bDragFromItself(false) { }

	virtual ~COleDragDropTabCtrl() { };


//private:
protected:

	// Overridables
	bool OnNewTabCtrlItems(int nInsertIndex, CSimpleArray<CTabItem>& items, IDataObject* pDataObject, DROPEFFECT& dropEffect)
	{
		return false;
	}


	bool OnDropTabCtrlItem(int nIndex, IDataObject *pDataObject, DROPEFFECT &dropEffect)
	{	// if return true, delete src items considering ctrl key
		return true;
	}


	void OnDeleteItemDrag(int nIndex) { }


	virtual HRESULT OnGetTabCtrlDataObject(CSimpleArray<int> arrIndex, IDataObject **ppDataObject) = 0;
#if 0
	{
		ATLASSERT(ppDataObject != NULL);
		*ppDataObject = NULL;
		return E_NOTIMPL;
	}
#endif

	// Mothods
	enum _hitTestFlag 
	{ 
		htInsetLeft, 	// �^�u�o�[�̍�
//		htInsetRight, 	// �^�u�o�[�̉E
		htItem, 		// �^�u
		htSeparator, 	// �^�u�̋��E
		htOutside, 		// �^�u�o�[�̗̈�O
		htWhole, 		// �^�u���Ȃ�?
		htOneself,		// �����̃^�u
		htOneselfLeft,	// �����̃^�u�̍�
		htOneselfRight,	// �����̃^�u�̉E
	};
	enum { kInset = 10 };

	// Drag���Ă���^�u�Ƃ��̍��E�ɃJ�[�\�������邩
	int _HitTestCurTabSide(CPoint point, _hitTestFlag &flag)
	{
		// Drag����Ă���^�u����Ȃ�
		if (m_arrCurDragItems.GetSize() == 1) {
			int nCurDragItemIndex = m_arrCurDragItems[0];

			/* Drag����Ă���^�u�̏�ɃJ�[�\�������邩�ǂ��� */
			if ( GetItem(nCurDragItemIndex).m_rcItem.PtInRect(point) ) {
				flag = htOneself;	// �����̃^�u�̏�Ƀ|�C���^������̂Ŗ����ɂ���
				return 0;
			}

			/* Drag����Ă���^�u�̍��̃^�u�̏�ɃJ�[�\�������邩�ǂ��� */
			if (_IsValidIndex(nCurDragItemIndex -1)) {
				const CRect &rcLeft = GetItem(nCurDragItemIndex -1).m_rcItem;
				if (rcLeft.PtInRect(point)) {
					// ���̃^�u��ɂ���
					flag = htOneselfLeft;
					return nCurDragItemIndex -1;
				}
			}

			/* Drag����Ă���^�u�̉E�̃^�u�̏�ɃJ�[�\�������邩�ǂ��� */
			if (_IsValidIndex(nCurDragItemIndex +1)) {
				const CRect &rcRight = GetItem(nCurDragItemIndex +1).m_rcItem;
				if (rcRight.PtInRect(point)) {
					// �E�̃^�u��ɂ���
					flag = htOneselfRight;
					return nCurDragItemIndex +1;
				}
			}
		}
		return -1;
	}

	// �ǂ̃^�u�̊ԂɃJ�[�\�������邩�ǂ���
	int _HitTestSeparator(CPoint point, vector<std::pair<int, CRect> >& vecRect, bool bInset = false)
	{
		DDTCTRACE( _T("COleDragDropTabCtrl::_HitTestSeparator\n") );

		int	cyGap = GetItemHeight();

		int	i = CTabBarConfig::s_bMultiLine ? 0 : m_nFirstIndexOnSingleLine;

		int	cyOffset = -1;
		for (; i < GetItemCount(); ++i) {
			const CRect &rcItem = GetItem(i).m_rcItem;
			if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
				continue;
			}

			/* ���[�̃^�u�̃C���f�b�N�X�Ƒ傫�����L������ */
			if (cyOffset != rcItem.top) {
				vecRect.push_back(std::make_pair(i, rcItem));
				cyOffset = rcItem.top;
			}

			// �^�u�̉E�[ 
			// |  �^�u   |�������̕���
			CRect rcSep( rcItem.right,
						 rcItem.top,
						 rcItem.right,
						 rcItem.top	  + cyGap);

			if (bInset) {
				rcSep.InflateRect(kInset, 0);
			} else {
				rcSep.InflateRect(rcItem.Width() / 2, 0);
			}

			if (rcSep.PtInRect(point)) {
				DDTCTRACE( _T(" �A�C�e���̊Ԃɂ���܂��I(%d)\n"), i );
				return i;
			}
		}

		DDTCTRACE( _T(" �A�C�e���̊Ԃɂ���܂���\n") );
		return -1;
	}

	// �^�u�o�[�̍�����ɃJ�[�\�������邩�ǂ���
	int _HitTestLeftSide(CPoint point, vector<std::pair<int, CRect> >& vecRect, bool bInset = false)
	{
		for (int i = 0; i < vecRect.size(); ++i) {
			if (bInset)
				vecRect[i].second.right = vecRect[i].second.left + kInset;
			vecRect[i].second.left = 0;			
			if (vecRect[i].second.PtInRect(point)) {
				return vecRect[i].first;
			}
		}
		return -1;
	}

	int _HitTestSideInset(CPoint point, int nInset = 0, bool *pbLeft = NULL)
	{
		DDTCTRACE( _T("COleDragDropTabCtrl::_HitTestSideInset\n") );
		
		int  cxGap			= (nInset == 0) ? (s_kcxGap * 2 + s_kcxSeparator) : nInset;
		int  cyGap			= GetItemHeight();

		bool bCheckLeftSide = CTabBarConfig::s_bMultiLine
							|| ( m_nFirstIndexOnSingleLine == 0 && !CTabBarConfig::s_bMultiLine );

		if (bCheckLeftSide && m_items.size() > 0) {
			int 		 nIndex = HitTest(point);
			DDTCTRACE( _T(" HitTest : %d\n"), nIndex );

			if (nIndex == -1) {
				return -1;
			}

			ATLASSERT( _IsValidIndex(nIndex) );
			CRect		 rc;
			GetClientRect(&rc);
			const CRect &rcItem = m_items[nIndex].m_rcItem;

			if (rcItem.left == s_kSideMargin) { 			
				// left side
				// �^�u�̍�������Hit����悤�ɕύX
				if ( CRect(rcItem.left, rcItem.top, s_kcxGap + (rcItem.Width() / 2), rcItem.top + cyGap).PtInRect(point) ) {
					if (pbLeft)
						*pbLeft = true;

					DDTCTRACE( _T(" �����ɂ���܂�(%d)\n"), nIndex );
					return nIndex;
				} else
					return -1;
			} else if (rcItem.right > rc.right - (rcItem.Width() / 2)) {	
				// right side
				// �E�[�̃^�u�̉E�����Ƀq�b�g���邩�ǂ���
				if ( CRect(rcItem.right - (rcItem.Width() / 2), rcItem.top, rc.right, rcItem.top + cyGap).PtInRect(point) ) {
					if (pbLeft)
						*pbLeft = false;

					DDTCTRACE( _T(" �E���ɂ���܂�(%d)\n"), nIndex );
					return nIndex;
				} else
					return -1;
			}
		}
		return -1;
	}

	int HitTestOnDragging(_hitTestFlag &flag, CPoint point, bool bInset = false)
	{
		int nIndex;

		// Drag����Ă���^�u���g�Ƃ��̍��E�ɃJ�[�\�������邩�ǂ���
		if (( nIndex = _HitTestCurTabSide(point, flag) ) != -1) {
			if (flag == htOneself) {
				return -1;	// Drag����Ă���̃^�u��ɃJ�[�\��������
			} else {
				return nIndex;
			}
		}

		vector<std::pair<int, CRect> > vecRect;

		// �^�u�̋��E��ɃJ�[�\�������邩�ǂ���
		if (( nIndex = _HitTestSeparator(point, vecRect, bInset) ) != -1) {
			// �^�u�̋��E��Hit����
			flag = htSeparator;
			return nIndex;
		}

		// �����ɃJ�[�\�������邩�ǂ���
		if (( nIndex = _HitTestLeftSide(point, vecRect, bInset) ) != -1) {
			flag = htInsetLeft;
			return nIndex;
		}

#if 0
		// inside of insets
		bool bLeft	= false;
		int  nIndex = _HitTestSideInset(point, s_nScrollInset, &bLeft);

		if (nIndex != -1) {
			// �^�u�o�[�̍��E�̒[��Hit����
			flag = bLeft ? htInsetLeft : htInsetRight;
			return nIndex;
		}

		// inside of the item rect
		if ( ( nIndex = HitTest(point) ) != -1 ) {		// hit test for item rect
			// �^�u��Drop���ꂽ
			flag = htItem;
			return nIndex;
		}
#endif
		// �ǂ̃^�u�O���[�v�͈̔͊O�ɑ��݂��邩
		{
			int cyOffset = 0;
			const int itemCount = GetItemCount();
			for (int i = 0; i < itemCount; ++i) {
				if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
					const CRect& rcItem = GetItem(i).m_rcItem;
					CRect rcGroup;
					rcGroup.top = cyOffset;
					rcGroup.left = 0;
					rcGroup.right = rcItem.right;
					rcGroup.bottom = rcItem.bottom;
					if (rcGroup.PtInRect(point)) {
						int index = i - 1;
						if (index == -1) {
							flag = htInsetLeft;
							return 0;
						} else {
							flag = htSeparator;
							return index;
						}
					}
					cyOffset = rcGroup.bottom;
				}
			}
		}


		// out side of items
		flag = htOutside;
		return -1;
	}


public:
	// Message map and handlers
	BEGIN_MSG_MAP(COleDragDropTabCtrl<T>)
		MESSAGE_HANDLER(WM_CREATE	, OnCreate)
		MESSAGE_HANDLER(WM_DESTROY	, OnDestroy)
		MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
		MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
		CHAIN_MSG_MAP( CTabCtrlImpl )
	END_MSG_MAP()


	// ���N���b�N
	LRESULT OnLButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		DDTCTRACE("COleDragDropTabCtrl::OnLButtonDown\n");
		POINT pt	 = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };

		if (wParam & MK_CONTROL) {
			bHandled = FALSE;
			return 0;
		}

		int   nIndex = HitTest(pt);

		if (nIndex != -1) {
			// �f�t�H���g�ł�Down���ɐ؂�ւ� release10��4
			// �ȑO��Up���ɐ؂�ւ���Down��D&D�̑O�i�K�Ƃ����d�l
			if (false/*GetTabCtrlExtendedStyle() & TAB2_EX_MOUSEDOWNSELECT*/) {
				// �^�u��؂�ւ���
				SetCurSel(nIndex, true);
				NMHDR nmhdr = { m_hWnd, (UINT_PTR)GetDlgCtrlID(), TCN_SELCHANGE };
				// �e�Ƀ^�u���؂�ւ�������Ƃ�ʒm����
				::SendMessage(GetParent(), WM_NOTIFY, (WPARAM) GetDlgCtrlID(), (LPARAM) &nmhdr);
			} else {
				_DoDragDrop(pt, (UINT) wParam, nIndex, true);
			}
		}

		return 0;
	}
	// �E�N���b�N
	LRESULT OnRButtonDown(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		DDTCTRACE("COleDragDropTabCtrl::OnRButtonDown\n");
		POINT pt	 = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };

		if (wParam & MK_CONTROL) {
			bHandled = FALSE;
			return 0;
		}

		int   nIndex = HitTest(pt);

		if (nIndex != -1) {
			_DoDragDrop(pt, (UINT) wParam, nIndex, false);
		}

		return 0;
	}

	LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		bHandled = FALSE;

		ATLASSERT( ::IsWindow(m_hWnd) );
		bool bReg = RegisterDragDrop();
		ATLASSERT(bReg == true);
		return 0;
	}
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL &bHandled)
	{
		bHandled = FALSE;
		RevokeDragDrop();
		return 0;
	}


public:
	/////////////////////////
	// IDropTarget

	// Overrides
	bool	OnScroll(UINT nScrollCode, UINT nPos, bool bDoScroll = true)
	{
		DDTCTRACE("COleDragDropTabCtrl::OnScroll\n");

		bool bResult = false;

		// �^�u�̕�����\��or�X�N���[���̕K�v���Ȃ���΋A��
		if ( !m_wndUpDown.IsWindowVisible() )
			return bResult;

		if (LOBYTE(nScrollCode) == SB_LINEUP) {
			if (bDoScroll)
				bResult = ScrollItem(false);
			else
				bResult = CanScrollItem(false);
		} else if (LOBYTE(nScrollCode) == SB_LINEDOWN) {
			if (bDoScroll)
				bResult = ScrollItem(true);
			else
				bResult = CanScrollItem(true);
		}

		return bResult;
	}

	void	SetDragImage(SHDRAGIMAGE& shdi)
	{
#if 0
		CClientDC		dc(m_hWnd);
		CBitmapHandle	bmp;
		int				nCurDragIndex = m_arrCurDragItems[0];
		CRect			rc			  = m_items[nCurDragIndex].m_rcItem;
		if (nCurDragIndex == GetCurSel()) {
			rc.left	-= s_kSideMargin;
			rc.right+= s_kSideMargin;
			bmp.CreateCompatibleBitmap(dc, rc.Width(), rc.Height());
		} else {
			rc.top	+= 2;
			bmp.CreateCompatibleBitmap(dc, rc.Width(), rc.Height());
		}

		CDC	memDC;
		memDC.CreateCompatibleDC(dc);
		memDC.SelectBitmap(bmp);
		memDC.BitBlt(0, 0, rc.Width(), rc.Height(), dc, rc.left, rc.right, SRCCOPY);

		bmp.Attach( AtlLoadBitmapImage(_T("G:\\�A�v���P�[�V����\\_\\Folsa-1.2\\src\\DiffStatus2.bmp")	, LR_LOADFROMFILE) );

		POINT	ptDrag;
		::GetCursorPos(&ptDrag);

		shdi.hbmpDragImage		= bmp;
		shdi.sizeDragImage.cx	= rc.Width();
		shdi.sizeDragImage.cy	= rc.Height();
		shdi.ptOffset.x			= ptDrag.x;
		shdi.ptOffset.y			= ptDrag.y;
		shdi.crColorKey			= CLR_NONE;
#endif
	}

	DROPEFFECT OnDragEnter(IDataObject *pDataObject, DWORD dwKeyState, CPoint point)
	{
		DWORD	dwEffect = _MtlStandardDropEffect(dwKeyState);

		/*
		CClientDC		dc(m_hWnd);
		CBitmapHandle	bmp;
		int				nCurDragIndex = m_arrCurDragItems[0];
		CRect			rc			  = m_items[nCurDragIndex].m_rcItem;
		if (nCurDragIndex == GetCurSel()) {
			rc.left	-= s_kSideMargin;
			rc.right+= s_kSideMargin;
			bmp.CreateCompatibleBitmap(dc, rc.Width(), rc.Height());
		} else {
			rc.top	+= 2;
			bmp.CreateCompatibleBitmap(dc, rc.Width(), rc.Height());
		}

		CDC	memDC;
		memDC.CreateCompatibleDC(dc);
		memDC.SelectBitmap(bmp);
		memDC.BitBlt(0, 0, rc.Width(), rc.Height(), dc, 0, 0, SRCCOPY);

//		m_DragImage.Create(rc.Width(), rc.Height(), ILC_COLORDDB, 1,10);
*/
		return dwEffect;
	}


	DROPEFFECT OnDragOver(IDataObject *pDataObject, DWORD dwKeyState, CPoint point, DROPEFFECT dropOkEffect)
	{
		if (m_bDragFromItself == false) {
			// �G�N�X�v���[���[����h���b�O���ꂽ�Ȃ�
//			int	nIndex = HitTest(point);

			return _MtlStandardDropEffect(dwKeyState);
		}

		_hitTestFlag flag;
		int 		 nIndex = HitTestOnDragging(flag, point);

		_DrawInsertionEdge(flag, nIndex);
#if 0
		if (flag == htItem && _IsSameIndexDropped(nIndex) )
			return DROPEFFECT_NONE;

		if (flag == htItem && m_bDragFromItself && m_arrCurDragItems.GetSize() > 1)
			return DROPEFFECT_NONE;
#endif
		if (!m_bDragFromItself)
			return _MtlStandardDropEffect(dwKeyState) | _MtlFollowDropEffect(dropOkEffect) | DROPEFFECT_COPY;

		return _MtlStandardDropEffect(dwKeyState);
	}


	DROPEFFECT OnDrop(IDataObject *pDataObject, DROPEFFECT dropEffect, DROPEFFECT dropEffectList, CPoint point)
	{
		_ClearInsertionEdge();

		if (m_bDragFromItself == false) {
			return dropEffect;		// CDonutTabBar�ŏ������ꂽ�̂ŋA��
		}

		T *pT	= static_cast<T *>(this);
		CSimpleArray<CTabItem> arrItems;

		_hitTestFlag	flag;
		int				nIndex = HitTestOnDragging(flag, point);

		if (flag == htInsetLeft) {
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// just move items
				MoveItems(nIndex, m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(nIndex, arrItems, pDataObject, dropEffect);

//				for (int i = 0; i < arrItems.GetSize(); ++i)
//					InsertItem(nIndex, arrItems[i]);
			}

		}
#if 0
		else if (flag == htInsetRight) {
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// just move items
				MoveItems(nIndex + 1, m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(nIndex + 1, arrItems, pDataObject, dropEffect);

				for (int i = 0; i < arrItems.GetSize(); ++i)
					InsertItem(nIndex + 1, arrItems[i]);
			}

		} 
#endif
		else if (flag == htOneselfLeft) {
			// �����̃^�u�̍��Ƀh���b�v���ꂽ
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// just move items
				MoveItems(nIndex, m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(nIndex + 1, arrItems, pDataObject, dropEffect);

//				for (int i = 0; i < arrItems.GetSize(); ++i)
//					InsertItem(nIndex + 1, arrItems[i]);
			}

		} else if (flag == htOneselfRight) {
			// �����̃^�u�̉E�Ƀh���b�v���ꂽ
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// just move items
				MoveItems(nIndex +1, m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(nIndex + 1, arrItems, pDataObject, dropEffect);

//				for (int i = 0; i < arrItems.GetSize(); ++i)
//					InsertItem(nIndex + 1, arrItems[i]);
			}

		}
#if 0
		else if ( flag == htItem && !_IsSameIndexDropped(nIndex) ) {
			// �^�u��Drop���ꂽ
			DROPEFFECT dropEffectPrev = dropEffect;
			bool	   bDelSrc		  = pT->OnDropTabCtrlItem(nIndex, pDataObject, dropEffect);

			if ( bDelSrc && m_bDragFromItself && (dropEffectPrev & DROPEFFECT_MOVE) ) {
				// just move items
				// do nothing
				for (int i = 0; i < m_arrCurDragItems.GetSize(); ++i)
//					pT->OnDeleteItemDrag(m_arrCurDragItems[i]);

				DeleteItems(m_arrCurDragItems);
			}

		}
#endif
		else if (flag == htSeparator) {
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// �^�u�̋��E��Hit�����̂�
				// nIndex + 1�ɑ}��
				MoveItems(nIndex + 1, m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(nIndex + 1, arrItems, pDataObject, dropEffect);

//				for (int i = 0; i < arrItems.GetSize(); ++i)
//					InsertItem(nIndex + 1, arrItems[i]);
			}

		} else if (flag == htOutside) {
			if ( m_bDragFromItself && (dropEffect & DROPEFFECT_MOVE) ) {
				// ��Ԍ��ɒǉ�����
				MoveItems(GetItemCount(), m_arrCurDragItems);
			} else {
//				pT->OnNewTabCtrlItems(GetItemCount(), arrItems, pDataObject, dropEffect);

//				for (int i = 0; i < arrItems.GetSize(); ++i)
//					AddItem(arrItems[i]);
			}
		}

		return dropEffect;
	}

	void OnDragLeave() { _ClearInsertionEdge(); }


//private:
protected:
	// Implementation

	void _ClearInsertionEdge()
	{
		if ( !m_rcInvalidateOnDrawingInsertionEdge.IsRectEmpty() ) {
			InvalidateRect(m_rcInvalidateOnDrawingInsertionEdge);
			UpdateWindow();
			m_rcInvalidateOnDrawingInsertionEdge.SetRectEmpty();
		}
	}

	// �C���T�[�g�}�[�N��`��
	bool _DrawInsertionEdge(_hitTestFlag flag, int nIndex)
	{
		CRect rcInvalidateOnDrawingInsertionEdge;

		if (flag == htOutside) {
			if (GetItemCount() > 0) {
				flag   = htSeparator;
				nIndex = GetItemCount() - 1;
			} else {
				flag = htWhole;
			}
		}

		CRect rcItem;

		if (flag == htInsetLeft) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem = m_items[nIndex].m_rcItem;
			_DrawInsertionEdgeAux(rcItem.TopLeft(), insertLeft);
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.left, rcItem.top, rcItem.left + s_kcxSeparator * 2, rcItem.bottom);

		}
#if 0
		else if (flag == htInsetRight) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem = m_items[nIndex].m_rcItem;
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.right - s_kcxSeparator * 2,
				rcItem.top,
				rcItem.right,
				rcItem.bottom);
			_DrawInsertionEdgeAux(CPoint(rcItem.right - s_kcxSeparator, rcItem.top), insertRight);

		}
#endif 
		else if (flag == htOneselfLeft) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem = m_items[nIndex].m_rcItem;
			CPoint ptSep(rcItem.left/* + s_kcxGap*/, rcItem.top);
			_DrawInsertionEdgeAux(ptSep, insertMiddle);
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.left - (s_kcxSeparator + 1),
				rcItem.top,
				rcItem.left + s_kcxGap * 2 - 1,
				rcItem.bottom);

		} else if (flag == htOneselfRight) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem = m_items[nIndex].m_rcItem;
			CPoint ptSep(rcItem.right/* + s_kcxGap*/, rcItem.top);
			_DrawInsertionEdgeAux(ptSep, insertMiddle);
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.right - (s_kcxSeparator + 1),
				rcItem.top,
				rcItem.right + s_kcxGap * 2 - 1,
				rcItem.bottom);

		} else if (flag == htSeparator) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem = m_items[nIndex].m_rcItem;
			CPoint ptSep(rcItem.right/* + s_kcxGap*/, rcItem.top);
			_DrawInsertionEdgeAux(ptSep, insertMiddle);
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.right - (s_kcxSeparator + 1),
				rcItem.top,
				rcItem.right + s_kcxGap * 2 - 1,
				rcItem.bottom);

		}
		else if (flag == htItem) {
			ATLASSERT( _IsValidIndex(nIndex) );
			rcItem	= m_items[nIndex].m_rcItem;
			_DrawInsertionEdgeAux(rcItem);
			rcInvalidateOnDrawingInsertionEdge = rcItem;

		}
		else if (flag == htWhole) {
			GetClientRect(rcItem);
			_DrawInsertionEdgeAux(rcItem.TopLeft(), insertLeft);
			rcInvalidateOnDrawingInsertionEdge.SetRect(
				rcItem.left, rcItem.top, rcItem.left + s_kcxSeparator * 2, rcItem.bottom);

		} else {
			_ClearInsertionEdge();
			return false;
		}

		if (rcInvalidateOnDrawingInsertionEdge != m_rcInvalidateOnDrawingInsertionEdge) {
			_ClearInsertionEdge();
			m_rcInvalidateOnDrawingInsertionEdge = rcInvalidateOnDrawingInsertionEdge;
		}

		return true;
	}


	enum insertFlags
	{ 
		insertLeft, 	// �^�u�o�[�̍�
		insertMiddle, 	// �^�u�̊�
		insertRight,	// �^�u�o�[�̉E
	};

	// I ��`�ʂ���(��ӁA��ӂ̒�����6�A���Ԃ̕���2�A�����̓^�u�̍����ɓ���)
	void _DrawInsertionEdgeAux(CPoint pt, insertFlags flag)
	{
		int 		 cy 	= GetItemHeight();
		int 		 sep	= s_kcxSeparator;
		CClientDC	 dc(m_hWnd);
		CBrush		 hbr;

		COLORREF InsertEdgeColor = CTabSkinDarkTheme::IsDarkMode() ? RGB(0xFF, 0xFF, 0xFF) : ::GetSysColor(COLOR_3DDKSHADOW);
		CPen penLine;
		penLine.CreatePen(BS_SOLID, 1, InsertEdgeColor);
		HPEN hpenOld = dc.SelectPen(penLine);

		hbr.CreateSolidBrush(InsertEdgeColor);
		dc.SetBrushOrg(pt.x, pt.y);
		CBrushHandle hbrOld = dc.SelectBrush(hbr);

		if (flag == insertLeft) {
			POINT pts[] = { { pt.x			, pt.y				  }, { pt.x 		 , pt.y + cy - 1 }, { pt.x + sep * 2 - 1, pt.y + cy - 1 },
							{ pt.x + sep - 1, pt.y + cy - sep - 1 }, { pt.x + sep - 1, pt.y + sep	 }, { pt.x + sep * 2 - 1, pt.y			}, };
			dc.Polygon( pts, _countof(pts) );

		} else if (flag == insertMiddle) {
			pt.x -= sep + 1;	// ������̍��W
			POINT pts[] = { { pt.x			, pt.y				  }, { pt.x + sep		 , pt.y + sep	 }, { pt.x + sep		, pt.y + cy - sep - 1 },
							{ pt.x			, pt.y + cy - 1 	  }, { pt.x + sep * 3 - 1, pt.y + cy - 1 }, { pt.x + sep * 2 - 1, pt.y + cy - sep - 1 },
							{ pt.x + sep * 2 - 1,  pt.y + sep	  }, { pt.x + sep * 3 - 1, pt.y 		 }	};
			dc.Polygon( pts, _countof(pts) );

		} else if (flag == insertRight) {
			POINT pts[] = { { pt.x - sep	,pt.y			}, { pt.x			, pt.y + sep	}, { pt.x			, pt.y + cy - sep - 1 },
							{ pt.x - sep	, pt.y + cy - 1 }, { pt.x + sep - 1 , pt.y + cy - 1 }, { pt.x + sep - 1 , pt.y				  }, };
			dc.Polygon( pts, _countof(pts) );
		}

		dc.SelectBrush(hbrOld);
		dc.SelectPen(hpenOld);
	}


	void _DrawInsertionEdgeAux(const CRect &rc)
	{
		CClientDC	 dc(m_hWnd);
		CBrush		 hbr;

		COLORREF InsertEdgeColor = CTabSkinDarkTheme::IsDarkMode() ? RGB(0xFF, 0xFF, 0xFF) : ::GetSysColor(COLOR_3DDKSHADOW);
		CPen penLine;
		penLine.CreatePen(BS_SOLID, 1, InsertEdgeColor);
		HPEN hpenOld = dc.SelectPen(penLine);

		hbr.CreateSolidBrush(InsertEdgeColor);
		dc.SetBrushOrg(rc.left, rc.top);
		CBrushHandle hbrOld = dc.SelectBrush(hbr);

		POINT		 pts[]	= { { rc.left	  , rc.top }, 
								{ rc.left	  , rc.bottom - 1 },
								{ rc.right - 1, rc.bottom - 1 },
								{ rc.right - 1, rc.top },
								{ rc.left	  , rc.top } 
							  };

		dc.Polyline( pts, _countof(pts) );

		dc.SelectBrush(hbrOld);
		dc.SelectPen(hpenOld);
	}


	bool _IsSameIndexDropped(int nDestIndex)
	{
		if (!m_bDragFromItself)
			return false;

		for (int i = 0; i < m_arrCurDragItems.GetSize(); ++i) {
			int nSrcIndex = m_arrCurDragItems[i];

			if (nSrcIndex == nDestIndex)
				return true;
		}

		return false;
	}

	virtual void OnVoidTabRemove(const CSimpleArray<int>& arrCurDragItems) = 0;

	void _DoDragDrop(CPoint pt, UINT nFlags, int nIndex, bool bLeftButton)
	{
		if ( PreDoDragDrop(m_hWnd, NULL, false) ) { 	// now dragging
			// Drop&Drag�J�n
			_HotItem(); 								// clean up hot item

			CComPtr<IDataObject> spDataObject;
			T * 	pT = static_cast<T *>(this);

			// set up current drag item index list
			GetCurMultiSelEx(m_arrCurDragItems, nIndex);

			HRESULT hr = pT->OnGetTabCtrlDataObject(m_arrCurDragItems, &spDataObject);
			if ( SUCCEEDED(hr) ) {	// Drag&Drop�J�n
				m_bDragFromItself = true;
//				DROPEFFECT dropEffect = DoDragDrop(spDataObject, DROPEFFECT_MOVE | DROPEFFECT_COPY);
				DROPEFFECT	dwEffectResult;
				hr = SHDoDragDrop(m_hWnd, spDataObject, NULL,  DROPEFFECT_MOVE | DROPEFFECT_COPY, &dwEffectResult);
				if (dwEffectResult == DROPEFFECT_NONE) {
					OnVoidTabRemove(m_arrCurDragItems);
				}

				m_bDragFromItself = false;

			}

			m_arrCurDragItems.RemoveAll();

		} else {
			// Drag&Drop�̑�������Ă��Ȃ������̂Ń^�u�؂�ւ�
			if (bLeftButton) {
				SetCurSel(nIndex, true);
				NMHDR nmhdr = { m_hWnd, (UINT_PTR)GetDlgCtrlID(), TCN_SELCHANGE };
				::SendMessage(GetParent(), WM_NOTIFY, (WPARAM) GetDlgCtrlID(), (LPARAM) &nmhdr);
			} else {
				SendMessage( WM_RBUTTONUP, (WPARAM) nFlags, MAKELPARAM(pt.x, pt.y) );
			}
		}
	}
};



////////////////////////////////////////////////////////////////////////////



}	//namespace MTL



using namespace MTL;

