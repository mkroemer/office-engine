import { createContext, useContext, useMemo, type ComponentType, type FormEvent, type MouseEvent, type ReactNode } from 'react'
import type { DropOption, IconName, Icons } from './edit/Toolbar'
import { IconsProvider } from './edit/Toolbar'
import { StringsProvider, type Locale, type StringsOverride } from './strings'

/**
 * A toolbar button. `onMouseDown` must reach the rendered element: it prevents the default so the document keeps
 * the focus and its selection while the button is clicked.
 */
export interface ButtonProps {
  title: string
  /** the icon, already resolved through the host's `icons` */
  icon?: ReactNode
  /** the icon's name, for a host that draws its own */
  iconName?: IconName
  /** text instead of (or after) the icon */
  label?: ReactNode
  /** content after the icon and label: a color swatch, a chevron */
  children?: ReactNode
  active?: boolean
  /** a popup this button opens is open */
  expanded?: boolean
  disabled?: boolean
  /** sized to its label rather than square */
  wide?: boolean
  onClick: () => void
  onMouseDown: (e: MouseEvent) => void
  /** the engine's own class names, for a host that keeps the engine's CSS for some buttons */
  className: string
}

/**
 * A dropdown: a value picker (`value` set: the chosen option is checked) or a command menu (`value` undefined).
 * A host's version must not move the focus out of the document while it is open, or editing commands lose their
 * selection: open it on mouse down with the default prevented, and handle its keys without focusing it.
 */
export interface SelectProps<T extends string | number = string | number> {
  title: string
  options: DropOption<T>[]
  onChange: (v: T) => void
  value?: T | ''
  placeholder?: string
  /** fixed trigger content (menus); null = icon only */
  label?: ReactNode
  /** Icon name, rendered by the host with `<Icon name={icon} />` (unlike Button's resolved node). */
  icon?: IconName
  width?: number
  menuWidth?: number
  disabled?: boolean
  /** popup aligned to the trigger's right edge */
  alignRight?: boolean
  /** extra class of the trigger button */
  className?: string
}

export interface DialogAction {
  label: string
  onClick?: () => void
  /** an auxiliary action, placed at the start of the built-in dialog's action row */
  extra?: boolean
  /** the main action: styled as such; with `submit`, also what Enter does */
  primary?: boolean
  /** submits the dialog's form (`onSubmit`); hosts with an external footer can call `onClick` instead */
  submit?: boolean
  autoFocus?: boolean
}

/** A modal dialog. Escape and a click on the backdrop call `onClose`; with `onSubmit` the body is a form. */
export interface DialogProps {
  title: string
  /** 'alert' for a question that needs an answer (role alertdialog) */
  kind?: 'dialog' | 'alert'
  onClose: () => void
  onSubmit?: () => void
  children: ReactNode
  actions: DialogAction[]
}

export interface ToastProps {
  text: string
}

/** The UI components a host can replace; each one left out keeps the engine's own. */
export interface UIComponents {
  Button: ComponentType<ButtonProps>
  Select: ComponentType<SelectProps>
  Separator: ComponentType
  Dialog: ComponentType<DialogProps>
  Toast: ComponentType<ToastProps>
}

const UIContext = createContext<Partial<UIComponents>>({})

/** Replaces UI components for everything below it; an outer provider's components stay for those left out. */
export function UIComponentsProvider({ components, children }: { components?: Partial<UIComponents>; children?: ReactNode }) {
  const outer = useContext(UIContext)
  const value = useMemo(() => (components ? { ...outer, ...components } : outer), [outer, components])
  return <UIContext.Provider value={value}>{children}</UIContext.Provider>
}

/** The host's components; undefined where the engine's own apply. */
export const useHostUI = (): Partial<UIComponents> => useContext(UIContext)

/** Everything a host customizes besides the theme, in one provider: for views mounted without `DocViewer`. */
export function OfficeUIProvider(props: { icons?: Icons; components?: Partial<UIComponents>; locale?: Locale | string; strings?: StringsOverride; children?: ReactNode }) {
  return (
    <StringsProvider locale={props.locale} strings={props.strings}>
      <IconsProvider icons={props.icons}>
        <UIComponentsProvider components={props.components}>{props.children}</UIComponentsProvider>
      </IconsProvider>
    </StringsProvider>
  )
}

// ---------------------------------------------------------------- the engine's own dialog and toast

export function DefaultDialog(props: DialogProps) {
  const body = (
    <>
      <div className="de-dialog-title">{props.title}</div>
      {props.children}
      <div className="de-dialog-btns">
        {props.actions.map((a, i) => (
          <button
            key={i}
            type={a.submit ? 'submit' : 'button'}
            className={'de-btn' + (a.primary ? ' de-btn-primary' : '') + (a.extra ? ' de-dialog-extra' : '')}
            autoFocus={a.autoFocus}
            onClick={a.submit && props.onSubmit ? undefined : a.onClick}
          >
            {a.label}
          </button>
        ))}
      </div>
    </>
  )
  const submit = (e: FormEvent) => {
    e.preventDefault()
    props.onSubmit?.()
  }
  return (
    <div
      className="de-dialog-back"
      onMouseDown={(e) => e.target === e.currentTarget && props.onClose()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          props.onClose()
        }
      }}
    >
      {props.onSubmit ? (
        <form className="de-dialog" role={props.kind === 'alert' ? 'alertdialog' : 'dialog'} aria-label={props.title} onSubmit={submit}>
          {body}
        </form>
      ) : (
        <div className="de-dialog" role={props.kind === 'alert' ? 'alertdialog' : 'dialog'} aria-label={props.title}>
          {body}
        </div>
      )}
    </div>
  )
}

export function DefaultToast({ text }: ToastProps) {
  return (
    <div className="de-toast" role="status">
      {text}
    </div>
  )
}
