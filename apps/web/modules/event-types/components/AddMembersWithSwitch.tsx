import AssignAllTeamMembers from "@calcom/features/eventtypes/components/AssignAllTeamMembers";
import type {
  CheckedSelectOption,
  CheckedTeamSelectCustomClassNames,
} from "@calcom/features/eventtypes/components/CheckedTeamSelect";
import CheckedTeamSelect from "@calcom/features/eventtypes/components/CheckedTeamSelect";
import type { Host, SettingsToggleClassNames, TeamMember } from "@calcom/features/eventtypes/lib/types";
import { useLocale } from "@calcom/lib/hooks/useLocale";
import { Label } from "@calcom/ui/components/form";
import { type ComponentProps, type Dispatch, type SetStateAction } from "react";
import type { Options } from "react-select";

interface IUserToValue {
  id: number | null;
  name: string | null;
  username: string | null;
  avatar: string;
  email: string;
  defaultScheduleId: number | null;
}

export const mapUserToValue = (
  { id, name, username, avatar, email, defaultScheduleId }: IUserToValue,
  pendingString: string
) => ({
  value: `${id || ""}`,
  label: `${name || email || ""}${!username ? ` (${pendingString})` : ""}`,
  avatar,
  email,
  defaultScheduleId,
});

const sortByLabel = (a: ReturnType<typeof mapUserToValue>, b: ReturnType<typeof mapUserToValue>) => {
  if (a.label < b.label) {
    return -1;
  }
  if (a.label > b.label) {
    return 1;
  }
  return 0;
};

const CheckedHostField = ({
  labelText,
  placeholder,
  options = [],
  isFixed,
  value,
  onChange,
  helperText,
  isRRWeightsEnabled,
  groupId,
  customClassNames,
  ...rest
}: {
  labelText?: string;
  placeholder: string;
  isFixed: boolean;
  value: Host[];
  onChange?: (options: Host[]) => void;
  options?: Options<CheckedSelectOption>;
  helperText?: React.ReactNode | string;
  isRRWeightsEnabled?: boolean;
  groupId: string | null;
} & Omit<Partial<ComponentProps<typeof CheckedTeamSelect>>, "onChange" | "value">) => {
  return (
    <div className="flex flex-col rounded-md">
      <div>
        {labelText ? <Label>{labelText}</Label> : <></>}
        <CheckedTeamSelect
          isOptionDisabled={(option) => !!value.find((host) => host.userId.toString() === option.value)}
          onChange={(options) => {
            onChange &&
              onChange(
                options.map((option) => ({
                  isFixed,
                  userId: parseInt(option.value, 10),
                  priority: option.priority ?? 2,
                  weight: option.weight ?? 100,
                  scheduleId: option.defaultScheduleId,
                  groupId: option.groupId,
                }))
              );
          }}
          value={(value || [])
            .filter(({ isFixed: _isFixed }) => isFixed === _isFixed)
            .reduce((acc, host) => {
              const option = options.find((member) => member.value === host.userId.toString());
              if (!option) return acc;

              acc.push({
                ...option,
                priority: host.priority ?? 2,
                isFixed,
                weight: host.weight ?? 100,
                groupId: host.groupId,
              });

              return acc;
            }, [] as CheckedSelectOption[])}
          controlShouldRenderValue={false}
          options={options}
          placeholder={placeholder}
          isRRWeightsEnabled={isRRWeightsEnabled}
          customClassNames={customClassNames}
          groupId={groupId}
          {...rest}
        />
      </div>
    </div>
  );
};

export type AddMembersWithSwitchCustomClassNames = {
  assingAllTeamMembers?: SettingsToggleClassNames;
  teamMemberSelect?: CheckedTeamSelectCustomClassNames;
};

export type AddMembersWithSwitchProps = {
  teamMembers: TeamMember[];
  value: Host[];
  onChange: (hosts: Host[]) => void;
  assignAllTeamMembers: boolean;
  setAssignAllTeamMembers: Dispatch<SetStateAction<boolean>>;
  automaticAddAllEnabled: boolean;
  onActive: () => void;
  isFixed: boolean;
  placeholder?: string;
  isRRWeightsEnabled?: boolean;
  teamId: number;
  groupId: string | null;
  "data-testid"?: string;
  customClassNames?: AddMembersWithSwitchCustomClassNames;
};

export function AddMembersWithSwitch({
  teamMembers,
  value,
  onChange,
  assignAllTeamMembers,
  setAssignAllTeamMembers,
  automaticAddAllEnabled,
  onActive,
  isFixed,
  placeholder = "",
  isRRWeightsEnabled,
  groupId,
  customClassNames,
  ...rest
}: AddMembersWithSwitchProps) {
  const { t } = useLocale();

  // When every team member is assigned there is no per-host list to render, so
  // the toggle is the whole control.
  if (assignAllTeamMembers) {
    return groupId ? null : (
      <AssignAllTeamMembers
        assignAllTeamMembers={assignAllTeamMembers}
        setAssignAllTeamMembers={setAssignAllTeamMembers}
        onActive={onActive}
        customClassNames={customClassNames?.assingAllTeamMembers}
      />
    );
  }

  return (
    <>
      <div className="mb-2">
        {automaticAddAllEnabled && !groupId && (
          <AssignAllTeamMembers
            assignAllTeamMembers={assignAllTeamMembers}
            setAssignAllTeamMembers={setAssignAllTeamMembers}
            onActive={onActive}
            customClassNames={customClassNames?.assingAllTeamMembers}
          />
        )}
      </div>
      <div className="mb-2">
        <CheckedHostField
          data-testid={rest["data-testid"]}
          value={value}
          onChange={onChange}
          isFixed={isFixed}
          className="mb-2"
          options={teamMembers
            .map((member) => ({
              ...member,
              groupId: groupId,
            }))
            .sort(sortByLabel)}
          placeholder={placeholder ?? t("add_attendees")}
          isRRWeightsEnabled={isRRWeightsEnabled}
          groupId={groupId}
          customClassNames={customClassNames?.teamMemberSelect}
        />
      </div>
    </>
  );
}

const AddMembersWithSwitchWrapper = ({
  containerClassName,
  ...props
}: AddMembersWithSwitchProps & {
  containerClassName?: string;
}) => {
  return (
    <div className="rounded-md">
      <div className={`flex flex-col rounded-md pb-2 pt-6 ${containerClassName}`}>
        <AddMembersWithSwitch {...props} />
      </div>
    </div>
  );
};

export default AddMembersWithSwitchWrapper;
