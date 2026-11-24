// Modified 2026-11-24 by janlewski for Vault Sheets.
import React, { useEffect, useState } from 'react'
import type { MenuProps, TabsProps } from 'antd'
import { Button, ConfigProvider, Dropdown, Flex, Splitter, Tabs, theme } from 'antd'
import { Notice } from 'obsidian'
import { TabType } from '../services/type'
import { useEventBus } from '../utils/useEventBus'
import { t } from '../lang/helpers'
import { log } from '../utils/log'
import { useEditorContext } from '../context/editorContext'
import { useUniver } from '../context/UniverContext'
import { useSheetStore } from '../context/SheetStoreProvider'
import { TAB_DEFAULT_ACTION, TAB_RENAME_ACTION, VIEW_REMOVE_ACTION } from '../services/reduce'
import { SheetTab } from './tabs/SheetTab'
import { RenameModal } from './components/RenameModal'
import { AppErrorBoundary } from './components/AppErrorBoundary'

function unavailableView(): React.ReactElement {
  return (
    <div className="vault-sheets-unavailable-view">
      {t('VIEW_TYPE_UNAVAILABLE')}
    </div>
  )
}

export const ContainerView = function ContainerView() {
  const { state, dispatch } = useSheetStore()
  const { univerApi } = useUniver()
  const { editor } = useEditorContext()
  const { plugin } = editor
  const [algorithm, setAlgorithm] = useState([]) // 设置主题
  const [triggerSource, setTriggerSource] = useState<string | null>(null) // 记录是点击哪个 tab 触发的下拉菜单
  const [renameModalVisible, setRenameModalVisible] = useState(false) // 重命名
  const [renameModalName, setRenameModalName] = useState('') // 重命名的名称

  const tabs = state.tabs.tabs || []
  const [activeKey, setActiveKey] = useState('sheet')

  const switchTab = () => {
    if (state.tabs.defaultActiveKey !== 'sheet') {
      setActiveKey(state.tabs.defaultActiveKey)
    }
  }
  const items: TabsProps['items'] = tabs.map((tab) => {
    if (tab.type === TabType.SHEET) {
      return {
        key: tab.key,
        label: tab.label,
        children: (
          <SheetTab
            key={tab.key}
            switchTab={switchTab}
          />
        ),
        forceRender: false,
      }
    }
    else if (tab.type === TabType.KANBAN || tab.type === TabType.CALENDAR || tab.type === TabType.DASHBOARD) {
      return {
        key: tab.key,
        label: tab.label,
        children: unavailableView(),
        forceRender: false,
      }
    }
    else {
      return {
        key: tab.key,
        label: tab.label,
        children: null,
        forceRender: true,
      }
    }
  })

  log('[ContainerView]', 'items', items)
  useEventBus<{ oldPath: string, newPath: string }>('fileRenamed', (payload) => {
    if (!payload) {
      return // 防止 undefined
    }
    const { oldPath, newPath } = payload
    const activeWorkbook = univerApi?.getActiveWorkbook()
    if (activeWorkbook && activeWorkbook.getName() === oldPath) {
      activeWorkbook.setName(newPath)
    }
  })

  useEffect(() => {
    if (!plugin) {
      return
    }
    const darkModal = plugin.settings.darkModal
    if (darkModal === 'dark') {
      setAlgorithm([theme.darkAlgorithm])
    }
    else {
      setAlgorithm([])
    }
  }, [plugin])

  const tabMenu: MenuProps['items'] = [
    {
      label: t('TAB_MENU_DEFAULT'),
      key: 'default',
    },
    {
      label: t('TAB_MENU_RENAME'),
      key: 'rename',
    },
    {
      label: t('TAB_MENU_DELETE'),
      key: 'delete',
    },
  ]

  const tabDropdownClick: MenuProps['onClick'] = (item) => {
    const { key } = item
    if (key === 'delete') {
      if (!triggerSource || triggerSource === 'sheet') {
        new Notice(t('CANNOT_DELETE_SHEET'))
        return
      }
      dispatch({ type: VIEW_REMOVE_ACTION, key: triggerSource })
      if (triggerSource == activeKey) {
        setActiveKey('sheet')
      }
      setTriggerSource(null)
    }
    if (key === 'default') {
      dispatch({ type: TAB_DEFAULT_ACTION, key: triggerSource || 'sheet' })
      setTriggerSource(null)
    }
    if (key === 'rename') {
      if (triggerSource) {
        const tab = tabs.find(t => t.key === triggerSource)
        if (tab) {
          setRenameModalName(tab.label)
          dispatch({ type: TAB_RENAME_ACTION, key: triggerSource, payload: tab.label })
        }
      }
      setRenameModalVisible(true)
    }
  }

  // 帮助菜单
  const helpTabMenu: MenuProps = {
    items: [
      {
        label: t('HELP_MENU_DOCS'),
        key: 'docs',
        onClick: () => {
          window.open('https://github.com/janlewski/vault-sheets#readme', '_blank')
        },
      },
    ],
  }

  // 自定义 TabBar 渲染函数
  const renderTabBar = (props: Record<string, unknown>, DefaultTabBar: React.ComponentType) => {
    return (
      <DefaultTabBar
        {...props}
        // 添加自定义的 tabBar 样式类名
        className="lj-tab-bar"
      >
        { (node: Record<string, unknown>) => {
          return (
            <div className="lj-tab-bar-item">
              <Dropdown
                menu={{
                  items: tabMenu,
                  onClick: tabDropdownClick,
                }}
                trigger={['contextMenu']}
                onOpenChange={
                  (open) => {
                    if (open) {
                      setTriggerSource(node.key)
                    }
                  }
                }
              >
                <Button
                  color="blue"
                  variant={activeKey === node.key ? 'solid' : 'outlined'}
                  onClick={() => {
                    setActiveKey(node.key)
                  }}
                >
                  {state.tabs.defaultActiveKey === node.key ? <span>★</span> : null}
                  {tabs.find(tab => tab.key === node.key)?.label || ''}
                </Button>
              </Dropdown>
            </div>
          )
        }}
      </DefaultTabBar>
    )
  }

  return (
    <AppErrorBoundary>
      <div className="lj-content-view" data-theme={plugin.settings.darkModal === 'dark' ? 'vault-sheets-dark' : ''}>
        <ConfigProvider
          theme={{
            algorithm,
            components: {
              Spin: {
                contentHeight: '100%',
              },
              Tabs: {
                horizontalMargin: '0',
              },
            },
          }}
        >
          <Splitter>
            <Splitter.Panel>
              <Tabs
                destroyOnHidden={false}
                className="full-height-tabs"
                size="small"
                type="card"
                items={items}
                activeKey={activeKey}
                renderTabBar={renderTabBar}
                tabBarExtraContent={{
                  right: (
                    <Flex gap="small">
                      <Dropdown menu={helpTabMenu} trigger={['hover']}>
                        <Button>?</Button>
                      </Dropdown>
                    </Flex>
                  ),
                }}
              />
            </Splitter.Panel>
          </Splitter>
          <RenameModal
            visible={renameModalVisible}
            name={renameModalName}
            onCancel={() => setRenameModalVisible(false)}
            onOk={(name) => {
              setRenameModalVisible(false)
              dispatch({
                type: TAB_RENAME_ACTION,
                key: triggerSource,
                payload: name,
              })
            }}
          />
        </ConfigProvider>
      </div>
    </AppErrorBoundary>
  )
}
